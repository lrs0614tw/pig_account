"""Fill the retained detail template without rebuilding its formatting."""
from copy import deepcopy
from datetime import datetime
from pathlib import Path
import re
from zipfile import ZipFile

from lxml import etree

from settlement_excel import roc_date, parse_pay_date

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = {'w': W}


def export_directory(root, today=None):
    directory = Path(root) / roc_date(today or datetime.today())
    directory.mkdir(parents=True, exist_ok=True)
    return directory


def available_path(directory, filename):
    path = Path(directory) / filename
    number = 2
    while path.exists():
        path = Path(directory) / f'{Path(filename).stem}_{number}{Path(filename).suffix}'
        number += 1
    return path


def detail_rows(farmers, account_info):
    rows = []
    for farmer in farmers:
        account = account_info(farmer.get('payee', farmer['farmer']), farmer['farmer'])
        rows.append((account['name'], account['acc'], str(farmer['heads']),
                     f"{farmer['net']:,}", account['note']))
    for farmer in farmers:
        if farmer['freight'] > 0:
            if farmer['driver'] in ('', '(無 / 自運)'):
                raise ValueError(f"{farmer['farmer']}有運費但未選擇司機。")
            account = account_info(farmer['driver'])
            rows.append((account['name'], account['acc'], '', f"{farmer['freight']:,}",
                         account['note'] or f"{farmer['farmer']}運費"))
    return rows


def _fill_cell(cell, value, prototype):
    """Keep paragraph/run/cell properties; change only text nodes."""
    texts = cell.findall('.//w:t', NS)
    if texts:
        texts[0].text = str(value)
        texts[0].set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        for text in texts[1:]:
            text.text = ''
        return
    if not value:
        return
    paragraph = cell.find('w:p', NS)
    if paragraph is None:
        raise ValueError('範本儲存格缺少段落。')
    # Blank template cells carry no runs. Reuse the corresponding populated
    # template cell's run style, while keeping the blank cell's paragraph style.
    run = etree.SubElement(paragraph, f'{{{W}}}r')
    properties = prototype.find('.//w:rPr', NS)
    if properties is not None:
        run.append(deepcopy(properties))
    text = etree.SubElement(run, f'{{{W}}}t')
    text.text = str(value)
    text.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')


def create_detail_document(template, target, farmers, account_info):
    if not farmers:
        raise ValueError('沒有可產出的明細。')
    template, target = Path(template), Path(target)
    if template.resolve() == target.resolve():
        raise ValueError('輸出不能覆蓋範本。')
    values = detail_rows(farmers, account_info)
    with ZipFile(template) as source:
        tree = etree.fromstring(source.read('word/document.xml'))
        table = tree.find('w:body/w:tbl', NS)
        if table is None:
            raise ValueError('價款明細表範本沒有表格。')
        rows = table.findall('w:tr', NS)
        fee_index = next((i for i, row in enumerate(rows)
                          if ''.join(row.findall('w:tc', NS)[0].itertext()).strip() == '手續費'), None)
        total_index = next((i for i, row in enumerate(rows)
                            if ''.join(row.findall('w:tc', NS)[0].itertext()).strip() == '合計'), None)
        if fee_index is None or total_index is None or fee_index < 2 or total_index <= fee_index:
            raise ValueError('範本需要明細列、手續費列與合計列。')
        prototype = deepcopy(rows[1])
        prototype_cells = prototype.findall('w:tc', NS)
        if len(prototype_cells) != 5:
            raise ValueError('範本明細必須有五欄。')
        detail = rows[1:fee_index]
        while len(detail) < len(values):
            row = deepcopy(prototype)
            rows[fee_index].addprevious(row)
            detail.append(row)
        for i, row in enumerate(detail):
            cells = row.findall('w:tc', NS)
            if len(cells) != 5:
                raise ValueError('範本明細列欄位數不一致。')
            for cell, value, style in zip(cells, values[i] if i < len(values) else ('',) * 5, prototype_cells):
                _fill_cell(cell, value, style)
        fees = sum(f['fa_fee'] for f in farmers)
        total = sum(f['net'] + f['freight'] + f['fa_fee'] for f in farmers)
        for row, data in [(rows[fee_index], ('手續費', '', '', f'{fees:,}', '')),
                          (rows[total_index], ('合計', '', str(sum(f['heads'] for f in farmers)), f'{total:,}', ''))]:
            for cell, value, style in zip(row.findall('w:tc', NS), data, prototype_cells):
                _fill_cell(cell, value, style)
        document_xml = etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)
        target.parent.mkdir(parents=True, exist_ok=True)
        with ZipFile(target, 'x') as output:
            for item in source.infolist():
                output.writestr(item, document_xml if item.filename == 'word/document.xml' else source.read(item.filename))
    return target


def _paragraph_text(paragraph):
    return ''.join(paragraph.xpath('.//w:t/text()', namespaces=NS))


def _replace_span(paragraph, start, end, value):
    """Replace a text span across runs without altering run properties."""
    offset = 0
    inserted = False
    for node in paragraph.findall('.//w:t', NS):
        text = node.text or ''
        node_end = offset + len(text)
        if node_end > start and offset < end:
            prefix = text[:max(0, start - offset)]
            suffix = text[max(0, end - offset):] if end < node_end else ''
            node.text = prefix + (str(value) if not inserted else '') + suffix
            node.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            inserted = True
        offset = node_end
    if not inserted:
        raise ValueError('範本文字位置無法替換。')


def _replace_match(paragraph, pattern, value):
    match = re.search(pattern, _paragraph_text(paragraph))
    if not match:
        raise ValueError(f'範本欄位不符預期：{_paragraph_text(paragraph)}')
    _replace_span(paragraph, match.start(), match.end(), value)


def _update_document_date(tree, pay_date):
    date = parse_pay_date(pay_date)
    matches = []
    pattern = r'民國\s*(\d+)年\s*(\d+)月\s*(\d+)日'
    for paragraph in tree.findall('w:body/w:p', NS):
        match = re.search(pattern, _paragraph_text(paragraph))
        if match:
            matches.append((paragraph, match))
    if len(matches) != 1:
        raise ValueError('範本必須有一個民國日期標題。')
    paragraph, match = matches[0]
    for group, number in reversed(list(enumerate((date.year - 1911, date.month, date.day), 1))):
        _replace_span(paragraph, match.start(group), match.end(group), str(number).zfill(len(match.group(group))))


def _write_patched_document(template, target, edit):
    template, target = Path(template), Path(target)
    if template.resolve() == target.resolve():
        raise ValueError('輸出不能覆蓋範本。')
    with ZipFile(template) as source:
        tree = etree.fromstring(source.read('word/document.xml'))
        edit(tree)
        xml = etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)
        target.parent.mkdir(parents=True, exist_ok=True)
        with ZipFile(target, 'x') as output:
            for item in source.infolist():
                output.writestr(item, xml if item.filename == 'word/document.xml' else source.read(item.filename))
    return target


def market_fee_template(farmers):
    markets = {f.get('market', '') for f in farmers}
    nantou = '南投縣市場' in markets
    changhua = '彰化縣市場' in markets
    if nantou and changhua:
        return '0.彰化南投手續費.docx'
    if nantou:
        return '0.南投手續費.docx'
    if changhua:
        return '0.彰化手續費.docx'
    return None


def create_market_fee_document(template, target, pay_date):
    return _write_patched_document(template, target, lambda tree: _update_document_date(tree, pay_date))


def create_expense_document(template, target, pay_date, farmers, voucher, chinese_amount):
    if not farmers:
        raise ValueError('沒有可產出的明細。')
    names = [f['farmer'] for f in farmers]
    payees = f'{names[0]}等{len(names)}人' if len(names) > 1 else names[0]
    heads = sum(f['heads'] for f in farmers)
    amount = sum(f['net'] + f['freight'] for f in farmers)
    purpose = f'付{payees}{heads}頭毛豬款。'

    def edit(tree):
        _update_document_date(tree, pay_date)
        tables = tree.findall('w:body/w:tbl', NS)
        if len(tables) < 2:
            raise ValueError('開支請示範本缺少主表。')
        table = tables[1]

        def after_label(label):
            for row in table.findall('w:tr', NS):
                cells = row.findall('w:tc', NS)
                for i, cell in enumerate(cells[:-1]):
                    if re.sub(r'\s+', '', ''.join(cell.xpath('.//w:t/text()', namespaces=NS))) == label:
                        paragraphs = cells[i + 1].findall('w:p', NS)
                        return paragraphs[0]
            raise ValueError(f'範本找不到「{label}」欄位。')

        _replace_match(after_label('採購商號或領款人'), r'.+', payees)
        _replace_match(after_label('新台幣'), r'[^:：\s].*元整', chinese_amount(amount))
        # Preserve the numeric amount's 16pt run; do not replace its 14pt prefix.
        _replace_match(after_label('開支金額'), r'[0-9][0-9,]*', f'{amount:,}')
        _replace_match(after_label('憑證'), r'[^\s]+(?=\s*張)', voucher)
        _replace_match(after_label('用途'), r'付.*毛豬款。', purpose)
        reasons = [p for p in table.findall('.//w:p', NS) if _paragraph_text(p).startswith('一、付')]
        if len(reasons) != 1:
            raise ValueError('範本找不到唯一的開支事由。')
        _replace_match(reasons[0], r'(?<=一、)付.*毛豬款。', purpose)

    return _write_patched_document(template, target, edit)


def create_collection_document(template, target, pay_date, farmers):
    """One input slip per page, retaining its member's original form geometry."""
    if not farmers:
        raise ValueError('沒有可產出的明細。')
    filled = parse_pay_date(pay_date)

    def edit(tree):
        body = tree.find('w:body', NS)
        elements = list(body)
        forms = {}
        for i, element in enumerate(elements):
            if element.tag == f'{{{W}}}p' and '毛豬共同運銷各項費用收取明細表' in _paragraph_text(element):
                block = elements[i:i + 4]
                if len(block) != 4 or block[1].tag != f'{{{W}}}tbl':
                    raise ValueError('費用明細表範本結構不符預期。')
                cells = block[1].find('w:tr', NS).findall('w:tc', NS)
                forms[_paragraph_text(cells[1]).strip()] = block
        if not forms:
            raise ValueError('範本沒有會員明細表。')
        driver_style = next(iter(forms.values()))[1].findall('w:tr', NS)[3].findall('w:tc', NS)[5]
        section = body.find('w:sectPr', NS)
        for element in list(body):
            if element is not section:
                body.remove(element)
        for index, farmer in enumerate(farmers):
            block = deepcopy(forms.get(farmer['farmer'], next(iter(forms.values()))))
            title, table, signature, contact = block
            # Explicit pagination replaces the source's multiple forms per page.
            properties = title.find('w:pPr', NS)
            if properties is None:
                properties = etree.Element(f'{{{W}}}pPr')
                title.insert(0, properties)
            page_break = properties.find('w:pageBreakBefore', NS)
            if page_break is None:
                page_break = etree.SubElement(properties, f'{{{W}}}pageBreakBefore')
            page_break.set(f'{{{W}}}val', '1' if index else '0')
            for element in block:
                for br in element.xpath('.//w:br[@w:type="page"]', namespaces=NS):
                    br.getparent().remove(br)
            rows = table.findall('w:tr', NS)
            header = rows[0].findall('w:tc', NS)
            values = rows[2].findall('w:tc', NS)
            if len(header) != 8 or len(values) != 8:
                raise ValueError('費用明細表欄位數不符預期。')
            auction = parse_pay_date(farmer['date'])
            market = farmer['market'].replace('市場', '肉品市場')
            for cell, value in [(header[1], farmer['farmer']),
                                (header[3], f'{auction.year - 1911}年{auction.month}月{auction.day}日'),
                                (header[5], market), (header[7], f"{farmer['heads']}頭")]:
                _fill_cell(cell, value, cell)
            # Clear every historical amount and note, including extra old slips.
            prototypes = deepcopy(values)
            for row in rows[2:]:
                for cell in row.findall('w:tc', NS):
                    _fill_cell(cell, '', cell)
            amounts = [farmer[key] for key in ('total', 'market_fee', 'accident', 'fa_fee', 'insurance', 'freight')]
            amounts += [0, farmer['net']]
            # Eight physical amount cells (merged grid columns are counted once).
            if len(values) != len(amounts):
                raise ValueError('費用金額欄位數不符預期。')
            for cell, value, prototype in zip(values, amounts, prototypes):
                _fill_cell(cell, f'{value:,}', prototype)
            driver = farmer.get('driver', '')
            if farmer['freight'] > 0 and driver not in ('', '(無 / 自運)'):
                note_cells = rows[3].findall('w:tc', NS)
                _fill_cell(note_cells[5], driver, driver_style)
            _replace_match(signature, r'\d+\.\d+\.\d+',
                           f'{filled.year - 1911}.{filled.month}.{filled.day}')
            for element in block:
                if section is not None:
                    section.addprevious(element)
                else:
                    body.append(element)

    return _write_patched_document(template, target, edit)
