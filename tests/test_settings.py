import os
import json
import tempfile
import unittest
from core.settings import Settings


class TestSettings(unittest.TestCase):
    """利用者設定 (core.settings) の単体テスト"""

    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.path = os.path.join(self.temp_dir.name, "sub", "settings.json")

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_defaults_when_file_missing(self):
        settings = Settings(self.path)
        self.assertEqual(settings.get("theme"), "Light")
        self.assertEqual(settings.get("last_csv_path"), "")
        self.assertTrue(settings.get("show_student_id"))

    def test_set_persists(self):
        Settings(self.path).set("theme", "Dark")
        self.assertEqual(Settings(self.path).get("theme"), "Dark")

    def test_broken_file_falls_back_to_defaults(self):
        os.makedirs(os.path.dirname(self.path))
        with open(self.path, "w", encoding="utf-8") as f:
            f.write("{ broken")
        self.assertEqual(Settings(self.path).get("theme"), "Light")

    def test_invalid_types_are_ignored(self):
        os.makedirs(os.path.dirname(self.path))
        with open(self.path, "w", encoding="utf-8") as f:
            json.dump({"theme": 1, "show_student_id": False}, f)
        settings = Settings(self.path)
        self.assertEqual(settings.get("theme"), "Light")
        self.assertFalse(settings.get("show_student_id"))


if __name__ == "__main__":
    unittest.main()
