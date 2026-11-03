import importlib.util
from pathlib import Path
import tkinter as tk
import sys
import unittest
from tempfile import TemporaryDirectory
from unittest.mock import patch

SOURCE = Path(__file__).resolve().parents[1] / '毛豬結算' / 'pig_accounting.py'
sys.path.insert(0, str(SOURCE.parent))
spec = importlib.util.spec_from_file_location('pig_accounting', SOURCE)
app_module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(app_module)


class InputTests(unittest.TestCase):
    def setUp(self):
        self.root = tk.Tk()
        self.root.withdraw()
        with patch.object(app_module.PigSettlementApp, '_read_cache', return_value={}):
            self.app = app_module.PigSettlementApp(self.root)

    def tearDown(self):
        self.root.destroy()

    def put(self, key, value):
        field = self.app.entries[key]
        field.delete(0, tk.END)
        field.insert(0, value)

    def fill(self, farmer='吳英吉', market='南投縣市場'):
        self.app.cbo_market.set(market)
        self.app.cbo_farmer.set(farmer)
        self.app.on_farmer_selected()
        self.put('heads', '10')
        self.put('total', '10500')
        self.app.auto_match_freight()

    def test_draft_folder_and_legacy_migration(self):
        import json
        with TemporaryDirectory() as directory:
            legacy = Path(directory) / 'pending_batches.json'
            target = Path(directory) / 'drafts' / 'pending_batches.json'
            old_data = {'old': {'farmers': [], 'auction_date': '115-11-03'}}
            legacy.write_text(json.dumps(old_data), encoding='utf-8')
            with patch.object(app_module, 'CACHE_FILE', str(target)), \
                 patch.object(app_module, 'LEGACY_CACHE_FILE', str(legacy)):
                self.assertEqual(self.app._read_cache(), old_data)
                self.assertTrue(target.exists())
                self.assertFalse(legacy.exists())
                new_data = dict(old_data, new={'farmers': []})
                self.app._write_cache(new_data)
                self.assertEqual(self.app._read_cache(), new_data)
                # An older file never replaces an already migrated draft store.
                legacy.write_text('{}', encoding='utf-8')
                self.assertEqual(self.app._read_cache(), new_data)
                self.assertTrue(legacy.exists())

    def test_strict_dates_and_individual_auction_dates(self):
        for value in ('1103', '115/11/03', '2026-11-03', '115-1-03', '115-11-3', '115-02-30', ' 115-11-03', ''):
            with self.assertRaises(ValueError, msg=value):
                app_module.parse_roc_input(value)
        self.assertEqual(app_module.roc_date(app_module.parse_roc_input('115-11-03')), '115-11-03')
        for date in ('115-10-31', '115-11-02'):
            self.fill()
            self.app.ent_auction_date.delete(0, tk.END)
            self.app.ent_auction_date.insert(0, date)
            self.app.add_farmer_to_list()
        self.app.ent_pay_date.delete(0, tk.END)
        self.app.ent_pay_date.insert(0, '115-11-05')
        with patch.object(self.app, '_read_cache', return_value={}), \
             patch.object(self.app, '_write_cache') as write, \
             patch.object(self.app, 'load_draft_list'), \
             patch.object(app_module.messagebox, 'showinfo'):
            self.assertTrue(self.app.save_as_draft())
            draft = next(iter(write.call_args.args[0].values()))
            self.assertNotIn('auction_date', draft)
            self.assertEqual(draft['pay_date'], '115-11-05')
            self.assertEqual([r['date'] for r in draft['farmers']], ['115-10-31', '115-11-02'])
        self.assertNotEqual(self.app.ent_auction_date.master, self.app.ent_pay_date.master)
        self.assertEqual([self.app.tree.item(row)['values'][0] for row in self.app.tree.get_children()], ['115-10-31', '115-11-02'])
        self.app.ent_pay_date.delete(0, tk.END)
        self.app.ent_pay_date.insert(0, '1105')
        with patch.object(app_module.messagebox, 'showwarning'), \
             patch.object(self.app, '_write_cache') as write, \
             patch.object(self.app, 'export_to_excel_sheet') as export:
            self.assertFalse(self.app.save_as_draft())
            self.app.finalize_batch()
            write.assert_not_called()
            export.assert_not_called()

    def test_legacy_draft_keeps_each_record_date(self):
        self.fill()
        self.app.add_farmer_to_list()
        first = self.app.current_farmers[0].copy()
        second = dict(first, date='2026-10-30')
        del first['date']
        draft = dict(auction_date='2026-10-31', pay_date='2026-11-05', farmers=[first, second])
        self.app.current_farmers = []
        self.app.draft_lookup = {'old': 'id'}
        self.app.cbo_drafts.set('old')
        with patch.object(self.app, '_read_cache', return_value={'id': draft}), \
             patch.object(app_module.messagebox, 'showinfo'):
            self.app.load_selected_draft()
        self.assertEqual(self.app.ent_pay_date.get(), '115-11-05')
        self.assertEqual([r['date'] for r in self.app.current_farmers], ['115-10-31', '115-10-30'])

    def test_sample_import_preserves_existing_drafts_and_labels(self):
        import json
        samples = json.loads((SOURCE.parents[1] / 'samples/測試草稿.json').read_text('utf-8'))
        first_id = next(iter(samples))
        existing = {'real': {'id': 'real', 'pay_date': '115-11-03', 'farmers': []},
                    first_id: dict(samples[first_id], name='【測試】已修改，請保留')}
        with patch.object(self.app, '_read_cache', return_value=existing), \
             patch.object(self.app, '_write_cache') as write, \
             patch.object(self.app, 'load_draft_list'), \
             patch.object(app_module.messagebox, 'showinfo'):
            self.app.import_sample_drafts()
            saved = write.call_args.args[0]
            self.assertEqual(len(saved), 6)
            self.assertEqual(saved[first_id]['name'], '【測試】已修改，請保留')
            self.assertIn('real', saved)
            write.reset_mock()
            self.app.import_sample_drafts()
            write.assert_not_called()
        with patch.object(self.app, '_read_cache', return_value=samples):
            self.app.load_draft_list()
            self.assertTrue(all('【測試】' in name for name in self.app.cbo_drafts['values']))
        self.app.active_draft_id = first_id
        self.app.current_farmers = samples[first_id]['farmers']
        with patch.object(self.app, '_read_cache', return_value=samples), \
             patch.object(self.app, '_write_cache') as write, \
             patch.object(self.app, 'load_draft_list'):
            self.app.save_as_draft(show_msg=False)
            self.assertEqual(write.call_args.args[0][first_id]['name'], samples[first_id]['name'])

    def test_numbers(self):
        for value, expected in [('1,000', 1000), ('1,000 + 200', 1200), ('', 0)]:
            self.put('total', value)
            self.assertEqual(self.app._get_int('total'), expected)
        for value in ('1,00', '100+abc', '100+', '-1', '1.5'):
            self.put('total', value)
            with self.assertRaises(ValueError):
                self.app._get_int('total')
            self.assertFalse(self.app.calculate_farmer_amounts())
            self.assertEqual(self.app.entries['net'].get(), '')

    def test_contracts_and_round(self):
        for farmer, market, driver, freight in [
            ('吳英吉', '南投縣市場', '杜虹儀', 550),
            ('張嘉顯', '新北市市場', '洪永原', 1200),
            ('吳櫂閎', '南投縣市場', '詹啓明', 9000),
        ]:
            self.fill(farmer, market)
            self.assertEqual(self.app.cbo_driver.get(), driver)
            self.assertEqual(self.app._get_int('freight'), freight)
            self.assertEqual(self.app.entries['fa_fee'].get(), '10')
        self.fill('張淑玲')
        self.assertEqual(self.app.cbo_driver.get(), '(無 / 自運)')
        self.assertEqual(self.app._get_int('freight'), 0)

    def test_payee_focus_preserves_selection(self):
        self.fill('張嘉顯')
        self.app.cbo_payee.set(app_module.PAYEE_PRESETS['張嘉顯'][1])
        self.app.on_farmer_selected()
        self.assertEqual(self.app.cbo_payee.get(), app_module.PAYEE_PRESETS['張嘉顯'][1])

    def test_edit_preserves_record_date_and_cancel(self):
        self.fill()
        self.app.ent_auction_date.delete(0, tk.END)
        self.app.ent_auction_date.insert(0, '115-11-01')
        self.app.add_farmer_to_list()
        original = self.app.current_farmers[0].copy()
        self.app.ent_auction_date.delete(0, tk.END)
        self.app.ent_auction_date.insert(0, '115-11-03')
        self.app.tree.selection_set(self.app.tree.get_children()[0])
        self.app.on_farmer_select(None)
        self.assertEqual(self.app.current_farmers, [original])
        self.assertEqual(self.app.ent_auction_date.get(), '115-11-01')
        self.put('accident', '100')
        with patch.object(app_module.messagebox, 'askyesno', return_value=True):
            self.app.cancel_input()
        self.assertEqual(self.app.current_farmers, [original])
        self.app.tree.selection_set(self.app.tree.get_children()[0])
        self.app.on_farmer_select(None)
        self.put('accident', '100')
        self.app.calculate_farmer_amounts()
        self.app.add_farmer_to_list()
        self.assertEqual(len(self.app.current_farmers), 1)
        self.assertEqual(self.app.current_farmers[0]['net'], original['net'] - 100)

    def test_pending_input_blocks_save_close_and_finalize(self):
        self.fill()
        with patch.object(app_module.messagebox, 'showwarning'), \
             patch.object(self.app, '_write_cache') as write, \
             patch.object(self.app, 'export_to_excel_sheet') as export, \
             patch.object(self.root, 'destroy') as destroy:
            self.app.save_as_draft()
            self.app.finalize_batch()
            self.app.on_closing()
            write.assert_not_called()
            export.assert_not_called()
            destroy.assert_not_called()

    def test_freight_without_driver_rejected(self):
        self.fill('張淑玲')
        self.put('freight', '100')
        with patch.object(app_module.messagebox, 'showwarning') as warning:
            self.app.ask_and_add_farmer()
            warning.assert_called_once()
        self.assertEqual(self.app.current_farmers, [])

    def test_detail_preview_does_not_finalize(self):
        self.fill()
        self.app.add_farmer_to_list()
        records = [record.copy() for record in self.app.current_farmers]
        with TemporaryDirectory() as directory, \
             patch.object(app_module, 'EXPORT_FOLDER', directory), \
             patch.object(app_module.win32.gencache, 'EnsureDispatch') as word, \
             patch.object(self.app, 'export_to_excel_sheet') as export, \
             patch.object(self.app, '_write_cache') as write:
            self.app.preview_detail_only()
            self.assertEqual(list(Path(directory).iterdir()), [])
            preview_path = Path(word.return_value.Documents.Open.call_args.args[0])
            self.assertFalse(preview_path.exists())
            word.return_value.Documents.Open.assert_called_once()
            export.assert_not_called()
            write.assert_not_called()
            self.assertEqual(self.app.current_farmers, records)

    def test_collection_preview_does_not_finalize(self):
        self.fill()
        self.app.add_farmer_to_list()
        records = [record.copy() for record in self.app.current_farmers]
        with TemporaryDirectory() as directory, \
             patch.object(app_module, 'EXPORT_FOLDER', directory), \
             patch.object(app_module.win32.gencache, 'EnsureDispatch') as word, \
             patch.object(self.app, 'export_to_excel_sheet') as export, \
             patch.object(self.app, '_write_cache') as write:
            self.app.preview_collection_only()
            self.assertEqual(list(Path(directory).iterdir()), [])
            preview_path = Path(word.return_value.Documents.Open.call_args.args[0])
            self.assertFalse(preview_path.exists())
            word.return_value.Documents.Open.assert_called_once()
            export.assert_not_called()
            write.assert_not_called()
            self.assertEqual(self.app.current_farmers, records)

    def test_finalize_keeps_monthly_excel_outside_date_folder(self):
        self.fill()
        self.app.ent_pay_date.delete(0, tk.END)
        self.app.ent_pay_date.insert(0, '115-11-25')
        self.app.add_farmer_to_list()
        with TemporaryDirectory() as directory:
            exports = Path(directory) / 'exports'
            exports.mkdir()
            master = exports / '豬款115.11.xlsm'
            master.write_bytes(b'monthly master unchanged by copying')
            exports = Path(directory) / 'exports'
            with patch.object(app_module, 'EXPORT_FOLDER', exports), \
                 patch.object(app_module.messagebox, 'askyesno', return_value=True), \
                 patch.object(app_module.messagebox, 'showinfo'), \
                 patch.object(app_module.simpledialog, 'askstring', return_value='1'), \
                 patch.object(self.app, 'export_to_excel_sheet', return_value=(master, '1103')), \
                 patch.object(self.app, 'preview_word_documents') as preview, \
                 patch.object(self.app, 'load_draft_list'):
                self.app.finalize_batch()
            snapshots = list(exports.rglob('*.xlsm'))
            self.assertEqual(len(snapshots), 1)
            self.assertEqual(snapshots[0], master)
            self.assertEqual(len(list(exports.rglob('1.*.docx'))), 1)
            self.assertEqual(len(list(exports.rglob('2.*.docx'))), 1)
            self.assertEqual(len(list(exports.rglob('3.*.docx'))), 1)
            self.assertEqual(len(list(exports.rglob('0.南投手續費.docx'))), 1)
            self.assertEqual({p.parent.name for p in exports.rglob('*.docx')}, {'115-11-25'})
            preview.assert_called_once()
            self.assertEqual(self.app.current_farmers, [])

    def test_monthly_export_location_and_legacy_move(self):
        with TemporaryDirectory() as directory:
            base = Path(directory) / 'app'
            base.mkdir()
            exports = Path(directory) / 'exports'
            legacy = base / '豬款115.11.xlsm'
            legacy.write_bytes(b'existing ledger')
            with patch.object(app_module, 'BASE_FOLDER', str(base)), \
                 patch.object(app_module, 'EXPORT_FOLDER', str(exports)), \
                 patch.object(app_module, 'export_monthly') as export:
                self.app.export_to_excel_sheet('115-11-03')
                self.assertEqual(export.call_args.args[0], str(exports))
                self.assertEqual((exports / legacy.name).read_bytes(), b'existing ledger')
                self.assertFalse(legacy.exists())
                legacy.write_bytes(b'conflicting ledger')
                with self.assertRaises(ValueError):
                    self.app.export_to_excel_sheet('115-11-03')
                self.assertEqual((exports / legacy.name).read_bytes(), b'existing ledger')
                self.assertEqual(legacy.read_bytes(), b'conflicting ledger')

    def test_expense_preview_does_not_finalize(self):
        self.fill()
        self.app.add_farmer_to_list()
        records = [record.copy() for record in self.app.current_farmers]
        with TemporaryDirectory() as directory, \
             patch.object(app_module, 'EXPORT_FOLDER', directory), \
             patch.object(app_module.simpledialog, 'askstring', return_value='3'), \
             patch.object(app_module.win32.gencache, 'EnsureDispatch') as word, \
             patch.object(self.app, 'export_to_excel_sheet') as export, \
             patch.object(self.app, '_write_cache') as write:
            self.app.preview_expenses_only()
            self.assertEqual(list(Path(directory).iterdir()), [])
            for call in word.return_value.Documents.Open.call_args_list:
                self.assertFalse(Path(call.args[0]).exists())
            self.assertEqual(word.return_value.Documents.Open.call_count, 2)
            export.assert_not_called()
            write.assert_not_called()
            self.assertEqual(self.app.current_farmers, records)

    def test_cancel_voucher_cancels_finalize(self):
        self.fill()
        self.app.add_farmer_to_list()
        with patch.object(app_module.messagebox, 'askyesno', return_value=True), \
             patch.object(app_module.simpledialog, 'askstring', return_value=None), \
             patch.object(self.app, 'export_to_excel_sheet') as export, \
             patch.object(self.app, 'create_detail_output') as detail:
            self.app.finalize_batch()
            export.assert_not_called()
            detail.assert_not_called()
            self.assertEqual(len(self.app.current_farmers), 1)


if __name__ == '__main__':
    unittest.main()
