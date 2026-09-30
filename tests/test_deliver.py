import tempfile
import unittest
import zipfile
from pathlib import Path

from deliver import package


class PackageTests(unittest.TestCase):
    def test_package_keeps_paths_and_filters_files(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "release"
            (source / "firmware").mkdir(parents=True)
            (source / ".git").mkdir()
            (source / "README.md").write_text("instructions", encoding="utf-8")
            (source / "firmware" / "app.bin").write_bytes(b"firmware")
            (source / "build.log").write_text("log", encoding="utf-8")
            (source / ".git" / "config").write_text("private", encoding="utf-8")
            output = root / "dist" / "release.zip"

            count = package(source, output, ["README.md", "firmware"], ["*.log"])

            self.assertEqual(count, 2)
            with zipfile.ZipFile(output) as archive:
                self.assertEqual(archive.namelist(), ["README.md", "firmware/app.bin"])
                self.assertEqual(archive.read("firmware/app.bin"), b"firmware")

    def test_missing_required_file_does_not_create_archive(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "release"
            source.mkdir()
            (source / "README.md").write_text("instructions", encoding="utf-8")
            output = root / "release.zip"

            with self.assertRaisesRegex(ValueError, "缺少必需文件"):
                package(source, output, ["firmware/app.bin"], [])
            self.assertFalse(output.exists())


if __name__ == "__main__":
    unittest.main()
