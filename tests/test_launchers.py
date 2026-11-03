import os
from pathlib import Path
import shutil
import subprocess
import sys
from tempfile import TemporaryDirectory
import unittest

ROOT = Path(__file__).resolve().parents[1]

@unittest.skipUnless(os.name == 'nt', 'Windows BAT launchers')
class LauncherTests(unittest.TestCase):
    def test_download_path_fallback_errors_and_logs(self):
        with TemporaryDirectory(prefix='pig_launcher_') as directory:
            package = Path(directory) / '測試 空格 & (下載)'
            (package / 'scripts').mkdir(parents=True)
            (package / 'app').mkdir()
            for name in ('啟動毛豬結算.bat', '安裝環境.bat'):
                shutil.copy2(ROOT / name, package / name)
            shutil.copy2(ROOT / 'scripts/launcher.ps1', package / 'scripts/launcher.ps1')
            env = {k: v for k, v in os.environ.items() if k.lower() not in ('path', 'pythonpath')}
            env['PATH'] = str(Path(sys.executable).parent) + ';' + os.environ['PATH']
            def run(name):
                return subprocess.run(['cmd.exe', '/d', '/c', 'call', str(package / name), '--no-pause'],
                                      cwd=directory, env=env, capture_output=True, timeout=60)
            for script, expected in [("print('LAUNCH_OK')", 0), ("raise RuntimeError('TEST_FAILURE')", 1)]:
                (package / 'app/pig_accounting.py').write_text(script, encoding='utf-8')
                result = run('啟動毛豬結算.bat')
                self.assertEqual(result.returncode, expected, repr(result.stdout + result.stderr))
            self.assertGreaterEqual(len(list((package / 'logs').glob('*.log'))), 2)
            logs = ''.join(p.read_text(encoding='utf-8-sig', errors='replace') for p in (package / 'logs').glob('*.log'))
            self.assertIn('LAUNCH_OK', logs)
            self.assertIn('TEST_FAILURE', logs)
            env['PATH'] = os.environ['SystemRoot'] + r'\System32'
            for name in ('啟動毛豬結算.bat', '安裝環境.bat'):
                self.assertEqual(run(name).returncode, 1)

    def test_bat_is_ascii_with_crlf(self):
        for name in ('啟動毛豬結算.bat', '安裝環境.bat'):
            data = (ROOT / name).read_bytes()
            data.decode('ascii')
            self.assertNotIn(b'\n', data.replace(b'\r\n', b''))
