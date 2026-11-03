from pathlib import Path
from tempfile import TemporaryDirectory
import sys
import unittest
from zipfile import ZipFile
from lxml import etree
ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / '毛豬結算'))
from settlement_documents import create_collection_document, NS
TEMPLATE = ROOT / '毛豬結算/templates/3.毛豬運銷各項費用收取明細表.docx'

def records():
    return [dict(farmer=name, date='115-11-02', market='南投縣市場', heads=50,
                 total=600000, market_fee=6000, accident=200, fa_fee=600,
                 insurance=100, freight=6000, net=587100, driver='洪永原')
            for name in ['吳英吉', '張嘉顯', '張嘉麟', '張盡固', '謝堎裕', '新會員', '吳英吉']]

class CollectionTests(unittest.TestCase):
    def test_per_record_pages_values_and_preserved_format(self):
        with TemporaryDirectory() as directory:
            target = create_collection_document(TEMPLATE, Path(directory)/'result.docx', '116-01-08', records())
            with ZipFile(TEMPLATE) as source, ZipFile(target) as output:
                for name in source.namelist():
                    if name != 'word/document.xml':
                        self.assertEqual(source.read(name), output.read(name))
                before = etree.fromstring(source.read('word/document.xml'))
                after = etree.fromstring(output.read('word/document.xml'))
            tables = after.findall('w:body/w:tbl', NS)
            self.assertEqual(len(tables), len(records()))
            self.assertEqual(len(after.xpath('//w:pageBreakBefore[@w:val="1"]', namespaces=NS)), 6)
            self.assertEqual(etree.tostring(before.find('w:body/w:sectPr', NS)), etree.tostring(after.find('w:body/w:sectPr', NS)))
            originals = before.findall('w:body/w:tbl', NS)
            for record, table, source_index in zip(records(), tables, [0,2,15,8,11,0,0]):
                text = ''.join(table.xpath('.//w:t/text()', namespaces=NS))
                self.assertIn(record['farmer'], text)
                self.assertIn('115年11月2日', text)
                self.assertIn('南投縣肉品市場', text)
                self.assertIn('587,100', text)
                self.assertIn('洪永原', text)
                self.assertNotIn('淘汰豬', text)
                self.assertNotIn('保費', text)
                for tag in ('w:tblPr', 'w:tblGrid'):
                    self.assertEqual(etree.tostring(table.find(tag, NS)), etree.tostring(originals[source_index].find(tag, NS)))
                self.assertEqual([etree.tostring(x) for x in table.findall('.//w:tcPr', NS)], [etree.tostring(x) for x in originals[source_index].findall('.//w:tcPr', NS)])
            text = ''.join(after.xpath('//w:t/text()', namespaces=NS))
            self.assertEqual(text.count('116.1.8'), len(records()))
            with self.assertRaises(FileExistsError):
                create_collection_document(TEMPLATE, target, '116-01-08', records())

if __name__ == '__main__':
    unittest.main()
