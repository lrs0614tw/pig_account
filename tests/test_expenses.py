from pathlib import Path
from tempfile import TemporaryDirectory
import sys
import unittest
from zipfile import ZipFile

from lxml import etree

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / '毛豬結算'))
from pig_accounting import num_to_chinese_upper
from settlement_documents import create_expense_document, create_market_fee_document, market_fee_template, NS

TEMPLATES = ROOT / '毛豬結算/templates'


class ExpenseTests(unittest.TestCase):
    def farmers(self):
        return [dict(farmer='張嘉顯', heads=10, net=123456, freight=1200, fa_fee=125,
                     market='南投縣市場'),
                dict(farmer='張淑玲', heads=5, net=20000, freight=0, fa_fee=20,
                     market='彰化縣市場')]

    def assert_format_preserved(self, template, output):
        with ZipFile(template) as source, ZipFile(output) as result:
            self.assertEqual(source.namelist(), result.namelist())
            for name in source.namelist():
                if name == 'word/document.xml':
                    before = etree.fromstring(source.read(name))
                    after = etree.fromstring(result.read(name))
                    for tree in (before, after):
                        for node in tree.findall('.//w:t', NS):
                            node.text = ''
                            node.attrib.pop('{http://www.w3.org/XML/1998/namespace}space', None)
                    self.assertEqual(etree.tostring(before), etree.tostring(after))
                else:
                    self.assertEqual(source.read(name), result.read(name), name)

    def test_market_selection(self):
        for markets, expected in [([], None), (['雲林縣市場'], None),
                                  (['南投縣市場', '南投縣市場'], '0.南投手續費.docx'),
                                  (['彰化縣市場'], '0.彰化手續費.docx'),
                                  (['彰化縣市場', '南投縣市場', '雲林縣市場'], '0.彰化南投手續費.docx')]:
            self.assertEqual(market_fee_template([{'market': value} for value in markets]), expected)

    def test_expense_values_and_full_format(self):
        template = TEMPLATES / '2.豬款開支請示.docx'
        with TemporaryDirectory() as directory:
            output = create_expense_document(template, Path(directory) / 'expense.docx',
                                             '116-01-08', self.farmers(), '參', num_to_chinese_upper)
            self.assert_format_preserved(template, output)
            with ZipFile(output) as z:
                tree = etree.fromstring(z.read('word/document.xml'))
                text = ''.join(tree.xpath('//w:t/text()', namespaces=NS))
                self.assertIn('民國 116年01月8日', text)
                self.assertIn('張嘉顯等2人', text)
                self.assertEqual(text.count('付張嘉顯等2人15頭毛豬款。'), 2)
                self.assertIn('144,656', text)
                self.assertIn(num_to_chinese_upper(144656), text)
                self.assertIn('參 張', text)
                self.assertNotIn('947,044', text)
                self.assertNotIn('陳尚成', text)
                self.assertNotIn('144,801', text)  # Fees must not be included.

    def test_fee_only_date_changes(self):
        with TemporaryDirectory() as directory:
            for name in ('0.南投手續費.docx', '0.彰化手續費.docx', '0.彰化南投手續費.docx'):
                template = TEMPLATES / name
                output = create_market_fee_document(template, Path(directory) / name, '116-01-08')
                self.assert_format_preserved(template, output)
                with ZipFile(template) as source, ZipFile(output) as result:
                    before = etree.fromstring(source.read('word/document.xml'))
                    after = etree.fromstring(result.read('word/document.xml'))
                    before_date = before.find('w:body/w:p[2]', NS)
                    after_date = after.find('w:body/w:p[2]', NS)
                    self.assertIn('民國116年01月08日', ''.join(after_date.itertext()))
                    before_date.getparent().remove(before_date)
                    after_date.getparent().remove(after_date)
                    self.assertEqual(etree.tostring(before), etree.tostring(after))


if __name__ == '__main__':
    unittest.main()
