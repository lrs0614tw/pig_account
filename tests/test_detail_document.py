from datetime import datetime
from hashlib import sha256
from pathlib import Path
from tempfile import TemporaryDirectory
import sys
import unittest
from zipfile import ZipFile

from lxml import etree

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / '毛豬結算'))
from pig_accounting import get_account_info
from settlement_documents import create_detail_document, detail_rows, export_directory, available_path, NS

TEMPLATE = ROOT / '毛豬結算/templates/1.毛豬運銷價款明細表.docx'


class DetailTests(unittest.TestCase):
    def farmer(self):
        return dict(farmer='張嘉顯', payee='曾淑美 (淘汰豬專用)', heads=10,
                    net=123456, freight=1200, driver='洪永原', fa_fee=125)

    def test_contract_accounts(self):
        rows = detail_rows([self.farmer()], get_account_info)
        self.assertEqual(rows[0], ('曾淑美', '1715-0', '10', '123,456', '張嘉顯汰豬專用'))
        self.assertEqual(rows[1][:4], ('洪永原', '6096-6', '', '1,200'))

    def test_template_preservation_and_amounts(self):
        digest = sha256(TEMPLATE.read_bytes()).hexdigest()
        with TemporaryDirectory() as directory:
            output = create_detail_document(TEMPLATE, Path(directory) / 'detail.docx',
                                            [self.farmer()], get_account_info)
            with ZipFile(TEMPLATE) as source, ZipFile(output) as result:
                self.assertEqual(source.namelist(), result.namelist())
                for name in source.namelist():
                    if name != 'word/document.xml':
                        self.assertEqual(source.read(name), result.read(name), name)
                before = etree.fromstring(source.read('word/document.xml'))
                after = etree.fromstring(result.read('word/document.xml'))
                for selector in ['w:body/w:sectPr', 'w:body/w:tbl/w:tblPr', 'w:body/w:tbl/w:tblGrid']:
                    self.assertEqual(etree.tostring(before.find(selector, NS)), etree.tostring(after.find(selector, NS)))
                original_rows = before.findall('w:body/w:tbl/w:tr', NS)
                rows = after.findall('w:body/w:tbl/w:tr', NS)
                self.assertEqual(len(rows), len(original_rows))
                for original, row in zip(original_rows, rows):
                    for selector in ['w:trPr', 'w:tc/w:tcPr', 'w:tc/w:p/w:pPr']:
                        self.assertEqual([etree.tostring(x) for x in original.findall(selector, NS)],
                                         [etree.tostring(x) for x in row.findall(selector, NS)])
                texts = lambda row: [''.join(c.xpath('.//w:t/text()', namespaces=NS)) for c in row.findall('w:tc', NS)]
                self.assertEqual(texts(rows[14]), ['手續費', '', '', '125', ''])
                self.assertEqual(texts(rows[15]), ['合計', '', '10', '124,781', ''])
                for row in rows[3:14]:
                    self.assertEqual(texts(row), [''] * 5)
            self.assertEqual(sha256(TEMPLATE.read_bytes()).hexdigest(), digest)

    def test_overflow_no_data_loss(self):
        with TemporaryDirectory() as directory:
            output = create_detail_document(TEMPLATE, Path(directory) / 'overflow.docx',
                                            [self.farmer()] * 8, get_account_info)
            with ZipFile(output) as result:
                tree = etree.fromstring(result.read('word/document.xml'))
                rows = tree.findall('w:body/w:tbl/w:tr', NS)
                self.assertEqual(len(rows), 19)  # header + 16 payments + fee + total
                self.assertIn('998,248', ''.join(rows[-1].itertext()))

    def test_date_folder_and_non_overwrite(self):
        with TemporaryDirectory() as directory:
            folder = export_directory(directory, datetime(2026, 11, 3))
            self.assertEqual(folder.name, '115-11-03')
            first = available_path(folder, 'detail.docx')
            first.write_bytes(b'original')
            self.assertEqual(available_path(folder, 'detail.docx').name, 'detail_2.docx')
            self.assertEqual(first.read_bytes(), b'original')


if __name__ == '__main__':
    unittest.main()
