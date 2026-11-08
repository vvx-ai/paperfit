import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

CLI=Path(__file__).resolve().parents[1]/'scripts'/'paperfit.py'

class PublicCLI(unittest.TestCase):
    def run_cli(self,*args):
        return subprocess.run([sys.executable,str(CLI),*map(str,args)],capture_output=True,text=True,cwd=tempfile.gettempdir())

    def test_all_commands_expose_help_and_version(self):
        sys.path.insert(0,str(CLI.parent))
        from cli import COMMANDS
        for command in COMMANDS:
            with self.subTest(command=command):
                self.assertEqual(self.run_cli(command,'--help').returncode,0)
        self.assertEqual(self.run_cli('--version').stdout.strip(),'PaperFit '+CLI.parent.parent.joinpath('VERSION').read_text().strip())

    def test_errors_have_nonzero_status_without_traceback(self):
        with tempfile.TemporaryDirectory() as td:
            result=self.run_cli('reviewed-template',Path(td)/'absent.docx',Path(td)/'missing.json',Path(td)/'out.docx')
            self.assertEqual(result.returncode,2)
            self.assertNotIn('Traceback',result.stderr)
            self.assertFalse((Path(td)/'out.docx').exists())
        self.assertEqual(self.run_cli('unknown-command').returncode,2)

    def test_public_pipeline_and_repeat_and_original_protection(self):
        with tempfile.TemporaryDirectory() as td:
            result=subprocess.run([sys.executable,str(CLI.with_name('smoke_test.py')),str(Path(td)/'smoke')],capture_output=True,text=True)
            self.assertEqual(result.returncode,0,result.stderr)
            self.assertEqual(json.loads(result.stdout)['status'],'passed')
