from contextlib import redirect_stderr, redirect_stdout
from io import StringIO
import json
import os
from pathlib import Path
import subprocess
import sys
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch
from uuid import uuid4

from core.run_lock import run_lock
from core.schema import AuditError
from run import ROOT, main

CLAIMS = (".claim-workpaper", ".claim-00_master_tb")
MOCK = ROOT / "data" / "mock"
ARGS = ["--adapter", "demo_kingdee", "--tb", str(MOCK / "mock_tb.xlsx"), "--gl", str(MOCK / "mock_gl.csv")]


class RunLockTest(unittest.TestCase):
    def setUp(self):
        temporary = TemporaryDirectory(dir=ROOT / "data" / "output")
        self.addCleanup(temporary.cleanup)
        self.output = Path(temporary.name)

    def child(self, claim, code):
        return subprocess.run(
            [sys.executable, "-c", "from pathlib import Path\nimport os, sys\n"
             "from core.run_lock import run_lock\nfrom core.schema import AuditError\n"
             "claim = Path(sys.argv[1])\n" + code, str(claim)],
            cwd=ROOT, timeout=10, capture_output=True, text=True, encoding="utf-8",
            env={**os.environ, "PYTHONUTF8": "1"},
        )

    def test_active_claim_rejects_another_process_and_keeps_owner_metadata(self):
        for name in CLAIMS:
            with self.subTest(name=name):
                claim = self.output / name
                with run_lock(claim, "busy"):
                    marker, = claim.iterdir()
                    data = marker.read_bytes()
                    owner = json.loads(data)
                    self.assertEqual(set(owner), {"version", "pid", "created_at_ns", "token", "hostname"})
                    self.assertEqual(owner["pid"], os.getpid())
                    self.assertEqual(marker.name, f'owner-{owner["token"]}.json')
                    child = self.child(claim, "try:\n"
                                              "    with run_lock(claim, 'busy'):\n"
                                              "        sys.exit(1)\n"
                                              "except AuditError as exc:\n"
                                              "    assert str(exc).startswith('busy'), str(exc)\n")
                    self.assertEqual(child.returncode, 0, child.stderr)
                    self.assertEqual(marker.read_bytes(), data)
                self.assertFalse(claim.exists())

    def test_cleans_own_claim_after_normal_and_exceptional_exit(self):
        for name in CLAIMS:
            claim = self.output / name
            with run_lock(claim, "busy"):
                self.assertTrue(claim.is_dir())
            self.assertFalse(claim.exists())
            for error in (AuditError, RuntimeError, KeyboardInterrupt):
                with self.subTest(name=name, error=error.__name__), self.assertRaisesRegex(error, "test error"):
                    with run_lock(claim, "busy"):
                        raise error("test error")
                self.assertFalse(claim.exists())

    def test_recovers_claim_after_own_child_exits_without_cleanup(self):
        for name in CLAIMS:
            with self.subTest(name=name):
                claim = self.output / name
                child = self.child(claim, "with run_lock(claim, 'busy'):\n    os._exit(0)\n")
                self.assertEqual(child.returncode, 0, child.stderr)
                previous, = claim.iterdir()
                with run_lock(claim, "busy"):
                    current, = claim.iterdir()
                    self.assertNotEqual(previous.name, current.name)
                    self.assertFalse(previous.exists())
                self.assertFalse(claim.exists())

    def test_unknown_claims_are_retained_and_report_manual_recovery(self):
        for state, name in zip(("old empty directory", "damaged metadata"), CLAIMS, strict=True):
            with self.subTest(state=state):
                claim = self.output / name
                claim.mkdir()
                if state == "damaged metadata":
                    (claim / f"owner-{uuid4().hex}.json").write_text("{", encoding="utf-8")
                before = tuple(claim.iterdir())
                with self.assertRaises(AuditError) as caught:
                    with run_lock(claim, "busy"):
                        self.fail("Unknown owner must block acquisition.")
                self.assertIn(str(claim), str(caught.exception))
                self.assertIn("确认相关任务均已停止后，手动移走", str(caught.exception))
                self.assertEqual(tuple(claim.iterdir()), before)

    def test_reclaims_neither_foreign_host_nor_unavailable_lock_state(self):
        claim = self.output / CLAIMS[0]
        child = self.child(claim, "with run_lock(claim, 'busy'):\n    os._exit(0)\n")
        self.assertEqual(child.returncode, 0, child.stderr)
        marker, = claim.iterdir()
        before = marker.read_bytes()
        with patch("core.run_lock.socket.gethostname", return_value=json.loads(before)["hostname"] + "-other"):
            with self.assertRaisesRegex(AuditError, "无法可靠确认运行锁状态"):
                with run_lock(claim, "busy"):
                    self.fail("A foreign host must block acquisition.")
        with patch("core.run_lock._try_lock", side_effect=OSError("Unsupported filesystem lock")):
            with self.assertRaisesRegex(AuditError, "无法可靠确认运行锁状态"):
                with run_lock(claim, "busy"):
                    self.fail("An unavailable lock probe must block acquisition.")
        self.assertEqual(marker.read_bytes(), before)

    def test_token_mismatch_preserves_owner_file(self):
        for name in CLAIMS:
            with self.subTest(name=name):
                claim = self.output / name
                with run_lock(claim, "busy"):
                    marker, = claim.iterdir()
                    owner = json.loads(marker.read_bytes())
                    owner["token"] = uuid4().hex
                    other = json.dumps(owner).encode("utf-8")
                    with marker.open("r+b") as stream:
                        stream.write(other)
                self.assertEqual(marker.read_bytes(), other)

    def test_both_cli_paths_reject_active_claims_before_starting_a_writer(self):
        for name, options, message in (
            (CLAIMS[0], ["--out", "workpaper.xlsx"], "同名底稿正在生成"),
            (CLAIMS[1], ["--subjects", "1122"], "另一个批量底稿任务正在生成"),
        ):
            with self.subTest(name=name), run_lock(self.output / name, "busy"):
                with (patch("run.store", return_value=self.output),
                      patch("run.workpaper_writer") as legacy, patch("run.ExcelWriter") as batch,
                      redirect_stdout(StringIO()) as stdout, redirect_stderr(StringIO()) as stderr):
                    self.assertEqual(main([*ARGS, *options]), 2)
                self.assertEqual(stdout.getvalue().strip(), "FAILED")
                self.assertIn(message, stderr.getvalue())
                self.assertIn(name, stderr.getvalue())
                legacy.assert_not_called()
                batch.assert_not_called()
        self.assertEqual(list(self.output.iterdir()), [])
