import tempfile
import unittest
from pathlib import Path

from docx import Document

from docreplacer_core import (
    case_replace,
    folder_rename_target,
    make_source_copy,
    replace_paragraph,
    name_from_logo,
)


class DocReplaceCoreTests(unittest.TestCase):
    def test_case_replace_preserves_case_style(self):
        self.assertEqual(case_replace("ACME acme Acme", "acme", "Zen Lodge"),
                         "ZEN LODGE zen lodge Zen Lodge")

    def test_replace_across_word_runs_keeps_surrounding_text(self):
        paragraph = Document().add_paragraph()
        first = paragraph.add_run("Welcome AC")
        first.bold = True
        middle = paragraph.add_run("ME Hotel")
        middle.italic = True
        last = paragraph.add_run(" today")

        changed = replace_paragraph(paragraph, "ACME Hotel", "Zen Lodge")

        self.assertTrue(changed)
        self.assertEqual(paragraph.text, "Welcome Zen Lodge today")
        self.assertTrue(first.bold)
        self.assertEqual(last.text, " today")

    def test_folder_target_uses_business_name_and_year(self):
        folder = Path("Documents - Old Hotel - 2026")
        self.assertEqual(
            folder_rename_target(folder, "New Hotel").name,
            "Documents - New Hotel - 2026",
        )

    def test_empty_source_is_rejected_before_copy(self):
        with tempfile.TemporaryDirectory() as temp:
            source = Path(temp) / "Documents - Old Hotel - 2026"
            source.mkdir()
            with self.assertRaises(FileNotFoundError):
                make_source_copy(source, "New Hotel")
            self.assertFalse((Path(temp) / "Documents - New Hotel - 2026").exists())

    def test_logo_filename_becomes_readable_business_name(self):
        self.assertEqual(
            name_from_logo(Path("zen_lapis-hotel.png")),
            "zen lapis hotel",
        )


if __name__ == "__main__":
    unittest.main()
