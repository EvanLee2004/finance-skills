import contextlib
import io
import json
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from skill_cli import dispatch
from test_eliminations import fixture, PERIOD

class InterfaceTests(unittest.TestCase):
    def test_describe(self):
        self.assertEqual(dispatch("describe")["cash_flow_method"], "out_both")

    def test_check_never_creates_output(self):
        _, evidence = fixture()
        with tempfile.TemporaryDirectory() as td:
            output = Path(td) / "new"
            request = {"parameters":{"period": PERIOD}, "output_dir":str(output)}
            result = dispatch("check", request, evidence)
            self.assertEqual(result["status"], "needs_input")
            self.assertIn("no_report_sources", result["preflight"]["issues"])
            self.assertFalse(output.exists())

    def test_changed_source_rejected(self):
        _, evidence = fixture()
        with tempfile.TemporaryDirectory() as td:
            source=Path(td)/"input.xlsx"
            source.write_bytes(b"fixture")
            request={"parameters":{"period": PERIOD}, "output_dir":str(Path(td)/"out"),
                     "files":{"reports":[{"local_path":str(source),"sha256":"wrong"}]}}
            self.assertIn("source_hash_mismatch:0", dispatch("check",request,evidence)["preflight"]["issues"])

    def test_run_json_and_true_complete(self):
        reports, evidence = fixture()
        with tempfile.TemporaryDirectory() as td:
            source=Path(td)/"input.xlsx"
            source.write_bytes(b"fixture")
            request={"parameters":{"period": PERIOD}, "output_dir":str(Path(td)/"out"),
                     "files":{"reports":[{"local_path":str(source),"name":"input.xlsx","file_id":"test"}]}}
            captured=io.StringIO()
            with patch("base_entry.parse_file", return_value=(reports, [])), contextlib.redirect_stdout(captured):
                result=dispatch("run",request,evidence)
            self.assertEqual(captured.getvalue(),"")
            self.assertEqual(result["status"],"complete")
            self.assertEqual(len(list((Path(td)/"out").glob("*.xlsx"))),3)
            json.dumps(result)

if __name__ == "__main__":
    unittest.main()
