"""Run the actual BAT/VBScript on Windows against temporary, fictitious INIs."""
from pathlib import Path
import os
import shutil
import subprocess
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
SECTION = "ODBC接続文字列"
A2 = "Ａ２テストＤＢ"
B7 = "Ｂ７テストＤＢ"


def fixture(first="OLD_A", second="OLD_B", newline="\r\n"):
    return newline.join([
        "[Other]", "Database=UNCHANGED", f"[{SECTION}]",
        ";A=Driver={SQL Server};Server=EXAMPLE_A;Database=COMMENT_A;",
        "  #B=Driver={SQL Server};Server=EXAMPLE_B;Database=COMMENT_B;",
        f"A=Driver={{SQL Server}};Server=EXAMPLE_A;Database={first};",
        f"B=Driver={{SQL Server}};Server=EXAMPLE_B;Database={second}",
        "[Other2]", "Database=UNCHANGED", "",
    ])


@unittest.skipUnless(os.name == "nt", "Windows cscript.exe is required")
class UpdaterTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="ini v5 ")
        self.directory = Path(self.temp.name)
        for name in ("set_ini_test.vbs", "set_A2.bat", "set_B7.bat", "set_test_7777DB.bat"):
            shutil.copyfile(ROOT / name, self.directory / name)
        self.ini = self.directory / "test.ini"
        self.original = fixture().encode("cp932")
        self.ini.write_bytes(self.original)
        self.env = dict(os.environ)
        self.env.update(INI_FILE=str(self.ini), NEW_VALUE=A2, INI_CHARSET="shift_jis",
                        TARGET_SECTION=SECTION, TARGET_KEY="Database", EXPECTED_MATCHES="2")
        self.cscript = str(Path(os.environ["SystemRoot"]) / "System32" / "cscript.exe")

    def tearDown(self):
        self.temp.cleanup()

    def run_vbs(self, *args):
        return subprocess.run([self.cscript, "//nologo", str(self.directory / "set_ini_test.vbs"), *args],
                              env=self.env, cwd=self.directory, capture_output=True, timeout=30)

    def run_bat(self, name, *args):
        command = name + " " + " ".join(args)
        return subprocess.run([os.environ.get("COMSPEC", "cmd.exe"), "/d", "/c", command],
                              env=self.env, cwd=self.directory, capture_output=True, timeout=30)

    def success(self, result):
        self.assertEqual(result.returncode, 0, (result.stdout + result.stderr).decode("cp932", errors="replace"))

    def test_environment_fullwidth_two_values(self):
        r = self.run_vbs("--env")
        self.success(r)
        self.assertIn(b"Updated: 2 value(s).", r.stdout)
        self.assertEqual(self.ini.read_bytes(), fixture(A2, A2).encode("cp932"))

    def test_a2_bat_cp932(self):
        self.success(self.run_bat("set_A2.bat", "/nopause"))
        self.assertEqual(self.ini.read_bytes(), fixture(A2, A2).encode("cp932"))

    def test_b7_bat_cp932(self):
        self.success(self.run_bat("set_B7.bat", "/nopause"))
        self.assertEqual(self.ini.read_bytes(), fixture(B7, B7).encode("cp932"))

    def test_original_bat_ascii_example(self):
        self.success(self.run_bat("set_test_7777DB.bat", "/nopause"))
        self.assertEqual(self.ini.read_bytes(), fixture("99DB", "99DB").encode("cp932"))

    def test_check_mode_writes_nothing(self):
        before = set(self.directory.iterdir())
        r = self.run_vbs("--check")
        self.success(r)
        self.assertIn(b"Would update: 2", r.stdout)
        self.assertEqual(self.ini.read_bytes(), self.original)
        self.assertEqual(set(self.directory.iterdir()), before)

    def test_bat_check_mode(self):
        self.success(self.run_bat("set_A2.bat", "/check", "/nopause"))
        self.assertEqual(self.ini.read_bytes(), self.original)

    def test_count_guard_no_partial_update(self):
        before = fixture().replace("Database=OLD_B", "Other=OLD_B").encode("cp932")
        self.ini.write_bytes(before)
        r = self.run_vbs("--env")
        self.assertEqual(r.returncode, 1)
        self.assertIn(b"E204:", r.stdout)
        self.assertEqual(self.ini.read_bytes(), before)

    def test_missing_key_is_named_error_and_no_write(self):
        self.env.pop("TARGET_KEY", None)
        r = self.run_vbs("--env")
        self.assertEqual(r.returncode, 1)
        self.assertIn(b"E014:", r.stdout)
        self.assertEqual(self.ini.read_bytes(), self.original)

    def test_unicode_escape_section(self):
        self.env["TARGET_SECTION"] = r"ODBC\u63A5\u7D9A\u6587\u5B57\u5217"
        self.success(self.run_vbs("--env"))
        self.assertEqual(self.ini.read_bytes(), fixture(A2, A2).encode("cp932"))

    def test_case_insensitive_database_key(self):
        self.ini.write_bytes(fixture().replace("Database=", "DATABASE=").encode("cp932"))
        self.env["TARGET_KEY"] = "database"
        self.success(self.run_vbs("--env"))
        self.assertEqual(self.ini.read_bytes(), fixture(A2, A2).replace("Database=", "DATABASE=").encode("cp932"))

    def test_already_equal(self):
        before = fixture(A2, A2).encode("cp932")
        self.ini.write_bytes(before)
        r = self.run_vbs("--env")
        self.success(r)
        self.assertIn(b"No change needed", r.stdout)
        self.assertEqual(self.ini.read_bytes(), before)

    def test_no_backup_or_temp_files_after_success(self):
        before = set(self.directory.iterdir())
        self.success(self.run_vbs("--env"))
        self.assertEqual(set(self.directory.iterdir()), before)

    def test_utf8_ini_supported(self):
        self.env["INI_CHARSET"] = "utf-8"
        self.ini.write_bytes(fixture().encode("utf-8"))
        self.success(self.run_vbs("--env"))
        self.assertEqual(self.ini.read_bytes(), fixture(A2, A2).encode("utf-8"))

    def test_legacy_five_arguments(self):
        self.success(self.run_vbs(str(self.ini), B7, "shift_jis", SECTION, "Database"))
        self.assertEqual(self.ini.read_bytes(), fixture(B7, B7).encode("cp932"))

    def test_line_endings_and_no_final_newline(self):
        self.ini.write_bytes(fixture(newline="\n").rstrip("\n").encode("cp932"))
        self.success(self.run_vbs("--env"))
        self.assertEqual(self.ini.read_bytes(), fixture(A2, A2, newline="\n").rstrip("\n").encode("cp932"))


if __name__ == "__main__":
    unittest.main(verbosity=2)
