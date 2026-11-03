"""Check GUI validation, report saving, and form construction."""

import tempfile
import tkinter as tk
import unittest
from pathlib import Path

from git import Actor, Repo
from openpyxl import load_workbook

from version_history_gui import VersionHistoryApp, generate_report, validate_inputs


class GuiTests(unittest.TestCase):
    """Verify the GUI entry point uses the same commit comparison engine."""

    def test_form(self) -> None:
        """Construct the four-field form without leaving a window open."""
        root = tk.Tk()
        root.withdraw()
        try:
            app = VersionHistoryApp(root)
            root.update_idletasks()
            self.assertEqual(len(app.controls), 7)
            self.assertFalse(app.busy)
        finally:
            root.destroy()

    def test_validation_and_report(self) -> None:
        """Save to a selected destination and preserve it on comparison error."""
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "tests.yml"
            output = Path(directory) / "chosen report.xlsx"
            with Repo.init(directory) as repo:
                actor = Actor("Test", "test@example.com")
                commits = []
                for version in ("1.2.3", "1.2.4"):
                    source.write_text(
                        f"version: {version}\nrequires:\n"
                        f"  # Compo Alignment\n  - asm/{version}@jenkins/l060\n",
                        encoding="utf-8",
                    )
                    repo.index.add(["tests.yml"])
                    commits.append(repo.index.commit(
                        version, author=actor, committer=actor,
                    ).hexsha)
                arguments = validate_inputs(
                    str(source), str(output), commits[0], commits[1],
                )
                self.assertEqual(generate_report(*arguments), 1)
                workbook = load_workbook(output)
                sheet = workbook["Version Changes"]
                self.assertEqual(sheet.cell(1, 3).value, "SP02b04")
                self.assertEqual(sheet.cell(2, 1).value, "Compo Alignment")
                workbook.close()
                original = output.read_bytes()
                with self.assertRaises(ValueError):
                    generate_report(source, output, "0" * 40, commits[1])
                self.assertEqual(output.read_bytes(), original)
                self.assertEqual(list(Path(directory).glob(".version_report_*")), [])
                with self.assertRaisesRegex(ValueError, "four fields"):
                    validate_inputs("", str(output), *commits)
                with self.assertRaisesRegex(ValueError, "end in .xlsx"):
                    validate_inputs(str(source), str(source), *commits)
                with self.assertRaisesRegex(ValueError, "hexadecimal"):
                    validate_inputs(str(source), str(output), "invalid", commits[1])


if __name__ == "__main__":
    unittest.main()
