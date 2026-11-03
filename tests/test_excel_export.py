from datetime import datetime
from copy import copy
import os
from pathlib import Path
import sys
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / '毛豬結算'))
from settlement_excel import parse_pay_date, monthly_path, next_sheet_name, export_monthly


class DateTests(unittest.TestCase):
    def test_dates(self):
        self.assertEqual(parse_pay_date('1103', datetime(2026, 1, 1)), datetime(2026, 11, 3))
        self.assertEqual(parse_pay_date('115-11-03'), datetime(2026, 11, 3))
        self.assertEqual(parse_pay_date('115/11/03'), datetime(2026, 11, 3))
        self.assertEqual(monthly_path(ROOT, parse_pay_date('2027-01-02')).name, '豬款116.01.xlsm')
        for value in ['0230', '1103abc', '2026-13-01']:
            with self.assertRaises(ValueError):
                parse_pay_date(value)

    def test_duplicate_names(self):
        self.assertEqual(next_sheet_name('1103', {'1103', '1103-2'}), '1103-3')


@unittest.skipUnless(os.environ.get('TEST_EXCEL') == '1', 'Native Excel integration test')
class NativeExcelTests(unittest.TestCase):
    def test_template_month_append_and_overflow(self):
        import openpyxl
        template = ROOT / '毛豬結算/templates/每日工作頁.xlsm'
        farmer = dict(farmer='測試', market='南投縣市場', date='2026-11-02', heads=10,
                      total=10500, market_fee=100, accident=0, fa_fee=10,
                      insurance=0, freight=550, driver='杜虹儀', net=9840)
        with tempfile.TemporaryDirectory() as folder:
            path, name = export_monthly(folder, template, '2026-11-03', [farmer], '2026-11-02')
            self.assertEqual(name, '1103')
            self.assertEqual(path.name, '豬款115.11.xlsm')
            _, second = export_monthly(folder, template, '2026-11-03', [farmer] * 7, '2026-11-02')
            self.assertEqual(second, '1103-2')
            book = openpyxl.load_workbook(path, data_only=True)
            original = openpyxl.load_workbook(template)
            sheet, overflow = book['1103'], book[second]
            self.assertEqual(sheet['O13'].value, 10390)
            self.assertEqual(sheet['O15'].value, 10400)
            self.assertEqual(overflow['O17'].value, 10390 * 7)
            self.assertEqual(overflow['O19'].value, 10400 * 7)
            self.assertEqual(overflow['O20'].value, 10400 * 7 - 15)
            for col in 'ABCDEFGHIJKLMNOP':
                self.assertEqual(sheet.column_dimensions[col].width,
                                 original.active.column_dimensions[col].width)
            for index in range(1, 19):
                self.assertEqual(sheet.row_dimensions[index].height,
                                 original.active.row_dimensions[index].height)
            for row in original.active.iter_rows(min_row=1, max_row=18, max_col=16):
                for cell in row:
                    self.assertEqual(copy(sheet[cell.coordinate].border), copy(cell.border), cell.coordinate)
                    self.assertEqual(copy(sheet[cell.coordinate].font), copy(cell.font), cell.coordinate)
            self.assertEqual(sheet.page_setup.orientation, 'landscape')
            self.assertEqual(sheet.page_setup.paperSize, original.active.page_setup.paperSize)
            self.assertEqual(sheet.page_margins, original.active.page_margins)
            # OOXML omits fitToWidth when it equals the default of one page.
            self.assertIn(sheet.page_setup.fitToWidth, (None, 1))
            self.assertTrue(sheet.sheet_properties.pageSetUpPr.fitToPage)
            self.assertIn('$A$1:$P$22', str(overflow.print_area))
            backups = list((Path(folder) / 'backup').glob('*.backup-*.xlsm'))
            self.assertEqual(len(backups), 1)
            self.assertFalse(list(Path(folder).glob('*.backup-*.xlsm')))
            backup_book = openpyxl.load_workbook(backups[0], data_only=True)
            self.assertEqual(backup_book.sheetnames, ['1103'])
            self.assertEqual(backup_book['1103']['O15'].value, 10400)
            backup_book.close()
            book.close()
            original.close()
            path2, _ = export_monthly(folder, template, '2027-01-02', [farmer], '2026-11-02')
            self.assertEqual(path2.name, '豬款116.01.xlsm')


if __name__ == '__main__':
    unittest.main()
