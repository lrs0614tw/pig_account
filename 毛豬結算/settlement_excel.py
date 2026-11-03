"""Monthly settlement files, using Excel's native worksheet copy."""
from datetime import datetime
from pathlib import Path
import os
import re
import shutil
import tempfile

import win32com.client as win32


def roc_date(date):
    return f'{date.year - 1911:03d}-{date.month:02d}-{date.day:02d}'


def display_date(value):
    try:
        return roc_date(parse_pay_date(value))
    except (ValueError, AttributeError):
        return value


def parse_roc_input(value):
    """Strict user input; legacy parsing remains available for saved records."""
    if not re.fullmatch(r'[0-9]{3}-[0-9]{2}-[0-9]{2}', value):
        raise ValueError('日期只接受民國 yyy-mm-dd，例如 115-11-03。')
    year, month, day = map(int, value.split('-'))
    try:
        if year < 1:
            raise ValueError
        return datetime(year + 1911, month, day)
    except ValueError as exc:
        raise ValueError('日期不存在，請輸入有效的民國 yyy-mm-dd 日期。') from exc


def parse_pay_date(value, today=None):
    today = today or datetime.today()
    value = value.strip()
    if len(value) == 4 and value.isascii() and value.isdigit():
        value = f"{today.year}-{value[:2]}-{value[2:]}"
    else:
        parts = value.replace('/', '-').split('-')
        if len(parts) == 3 and parts[0].isdigit() and len(parts[0]) <= 3:
            value = f'{int(parts[0]) + 1911}-{parts[1]}-{parts[2]}'
    try:
        return datetime.strptime(value, "%Y-%m-%d")
    except ValueError as exc:
        raise ValueError("日期請輸入民國年，例如 115-11-03；出帳日也可填 1103。") from exc


def monthly_path(folder, pay_date):
    return Path(folder) / f"豬款{pay_date.year - 1911}.{pay_date.month:02d}.xlsm"


def payment_rows(farmers, auction_date):
    rows = []
    markets = {name: index for index, name in enumerate(
        ['新北市市場', '桃園縣市場', '南投縣市場', '彰化縣市場', '雲林縣市場'])}
    for farmer in farmers:
        values = [None] * 16
        market = farmer.get('market', '')
        if market not in markets:
            raise ValueError(f"{farmer['farmer']}尚未設定有效市場，請先修改該筆明細。")
        values[markets[market]] = '✓'
        date = parse_pay_date(farmer.get('date', auction_date))
        # Excel serial date avoids locale-dependent date strings.
        values[5] = (date - datetime(1899, 12, 30)).days
        values[6:15] = [farmer[key] for key in
                       ('farmer', 'heads', 'total', 'market_fee', 'accident',
                        'fa_fee', 'insurance', 'freight', 'net')]
        rows.append(values)
    for farmer in farmers:
        if farmer['freight'] > 0:
            if farmer['driver'] in ('', '(無 / 自運)'):
                raise ValueError(f"{farmer['farmer']}有運費但未選擇司機。")
            values = [None] * 16
            values[6] = farmer['driver']
            values[14] = farmer['freight']
            values[15] = f"{farmer['farmer']}運費"
            rows.append(values)
    return rows


def next_sheet_name(base, existing):
    if base not in existing:
        return base
    number = 2
    while f'{base}-{number}' in existing:
        number += 1
    return f'{base}-{number}'


def export_monthly(folder, template, pay_date, farmers, auction_date):
    """Copy the native template and save atomically; never replace an existing day."""
    date = parse_pay_date(pay_date)
    target = monthly_path(folder, date)
    template = Path(template).resolve()
    if not template.is_file():
        raise FileNotFoundError(f'找不到每日工作頁範本：{template}')
    monthly_template = template.with_name('月報表.xlsm')
    if not monthly_template.is_file():
        raise FileNotFoundError(f'找不到月報表範本：{monthly_template}')
    rows = payment_rows(farmers, auction_date)
    if not rows:
        raise ValueError('沒有可匯出的明細。')
    target.parent.mkdir(parents=True, exist_ok=True)
    initial_stat = target.stat() if target.exists() else None
    with tempfile.TemporaryDirectory(prefix='settlement_') as temp_dir:
        staging = Path(temp_dir) / target.name
        local_template = Path(temp_dir) / 'daily_template.xlsm'
        shutil.copy2(template, local_template)
        local_monthly = Path(temp_dir) / 'monthly_template.xlsm'
        shutil.copy2(monthly_template, local_monthly)
        if initial_stat:
            shutil.copy2(target, staging)
        excel = win32.gencache.EnsureDispatch(win32.DispatchEx('Excel.Application'))
        workbook = source = monthly_source = None
        try:
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.AutomationSecurity = 3  # Do not execute workbook macros.
            source = excel.Workbooks.Open(str(local_template), UpdateLinks=0, ReadOnly=True)
            if initial_stat:
                workbook = excel.Workbooks.Open(str(staging.resolve()), UpdateLinks=0)
                existing = {sheet.Name for sheet in workbook.Worksheets}
                source.Worksheets(1).Copy(After=workbook.Worksheets(workbook.Worksheets.Count))
                sheet = next(item for item in workbook.Worksheets if item.Name not in existing)
            else:
                existing = set()
                source.Worksheets(1).Copy()
                workbook = excel.ActiveWorkbook
                sheet = workbook.Worksheets(1)
            sheet_name = next_sheet_name(date.strftime('%m%d'), existing)
            sheet.Name = sheet_name
            extra = max(0, len(rows) - 10)
            if extra:
                sheet.Rows(f'12:{11 + extra}').Insert()
                sheet.Rows(9).Copy()
                sheet.Rows(f'12:{11 + extra}').PasteSpecial(Paste=-4122)  # Formats only.
                sheet.Rows(f'12:{11 + extra}').RowHeight = sheet.Rows(9).RowHeight
                excel.CutCopyMode = False
            last_detail = 12 + extra
            subtotal, total = 13 + extra, 15 + extra
            sheet.Range(f'A3:Q{last_detail}').ClearContents()
            last_date = None
            for index, values in enumerate(rows, 3):
                values = list(values)
                if values[5] is not None:
                    if values[5] == last_date:
                        values[5] = None
                    else:
                        last_date = values[5]
                sheet.Range(f'A{index}:P{index}').Value = (tuple(values),)
                if values[5] is not None:
                    sheet.Cells(index, 6).NumberFormat = sheet.Cells(3, 6).NumberFormat
                # Keep template heights unless wrapped text needs more room.
                minimum_height = sheet.Rows(index).RowHeight
                sheet.Rows(index).AutoFit()
                sheet.Rows(index).RowHeight = max(minimum_height, sheet.Rows(index).RowHeight)
            sheet.Cells(subtotal, 15).Formula = f'=SUM(O3:O{last_detail})'
            # Use Python-calculated fees from the input, preserving Python round().
            for column in 'HIJKLM':
                sheet.Range(f'{column}{total}').Formula = f'=SUM({column}3:{column}{last_detail})'
            sheet.Cells(total, 15).Formula = f'=O{subtotal}+L{total}'
            # Existing template deductions and labels remain in the copied sheet.
            for index in list(range(3, subtotal + 1)) + list(range(total, total + 4)):
                sheet.Cells(index, 19).Formula = f'=SUBSTITUTE(NUMBERSTRING(O{index}, 2), "零", "")'
            sheet.PageSetup.PrintArea = f'$A$1:$P${total + 3}'
            sheet.PageSetup.Zoom = False
            sheet.PageSetup.FitToPagesWide = 1
            sheet.PageSetup.FitToPagesTall = False
            sheet.PageSetup.PrintTitleRows = '$1:$2'
            # Copy only the report sheet, never historical daily sheets in a template.
            if '月報表' not in {item.Name for item in workbook.Worksheets}:
                monthly_source = excel.Workbooks.Open(str(local_monthly), UpdateLinks=0, ReadOnly=True)
                monthly_source.Worksheets('月報表').Copy(Before=workbook.Worksheets(1))
            report = workbook.Worksheets('月報表')
            if report.Index != 1:
                report.Move(Before=workbook.Worksheets(1))
            report.Range('A1').Value = f'土庫鎮農會{date.year - 1911}年{date.month}月份供銷部業務月報表：'
            report.Range('C2:D2').Value = ((f'{date.year - 1911}年{date.month}月',) * 2,)
            references = {column: [] for column in ('H', 'I', 'L')}
            for daily in workbook.Worksheets:
                if not re.fullmatch(date.strftime('%m') + r'\d{2}(?:-\d+)?', daily.Name):
                    continue
                # Locate the footer so inserted overflow rows are included.
                labels = daily.Range(f'P1:P{daily.UsedRange.Row + daily.UsedRange.Rows.Count - 1}').Value
                totals = [i for i, value in enumerate(labels, 1) if value[0] == '總金額']
                if len(totals) != 1 or totals[0] < 6:
                    raise ValueError(f'工作頁「{daily.Name}」找不到唯一的總金額列，未更新帳本。')
                last_row = totals[0] - 3
                for column in references:
                    references[column].append(f"'{daily.Name}'!{column}3:{column}{last_row}")
            for address, column in [('C11', 'H'), ('D11', 'I'), ('F11', 'L')]:
                report.Range(address).Formula = '=SUM(' + ','.join(references[column]) + ')'
            report.Activate()
            excel.Calculate()
            workbook.SaveAs(str(staging.resolve()), FileFormat=52)
        finally:
            try:
                if workbook is not None:
                    workbook.Close(SaveChanges=False)
                if source is not None:
                    source.Close(SaveChanges=False)
                if monthly_source is not None:
                    monthly_source.Close(SaveChanges=False)
            finally:
                excel.Quit()
        if initial_stat:
            latest = target.stat()
            if (latest.st_size, latest.st_mtime_ns) != (initial_stat.st_size, initial_stat.st_mtime_ns):
                raise RuntimeError('月份檔案在匯出期間有變動，請關閉 Excel 後重試；未覆寫原檔。')
            backup_folder = target.parent / 'backup'
            backup_folder.mkdir(exist_ok=True)
            backup = backup_folder / f'{target.stem}.backup-{datetime.now():%Y%m%d-%H%M%S-%f}.xlsm'
            shutil.copy2(target, backup)
            os.replace(staging, target)
        else:
            # Windows rename refuses to overwrite a file created concurrently.
            os.rename(staging, target)
    return target, sheet_name
