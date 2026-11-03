"""Integration checks using temporary repositories and workbooks."""

import tempfile
import unittest
from pathlib import Path

from git import Actor, Repo
from openpyxl import load_workbook

from version_history import COLUMNS, find_changes, version_at, write_report


class HistoryTests(unittest.TestCase):
    """Exercise real Git snapshots and the generated Excel report."""

    def test_history_and_workbooks(self) -> None:
        """Find transitions on branches and retain safe cells and headers."""
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            with Repo.init(root) as repo:
                actor = Actor("Test Author", "author@example.com")

                def commit(contents: str, message: str):
                    (root / "tests.yml").write_text(contents, encoding="utf-8")
                    repo.index.add(["tests.yml"])
                    return repo.index.commit(
                        message, author=actor, committer=actor,
                    )

                initial = commit(
                    "version: 2.1.2\nrequires:\n- Library/1.0.0\n",
                    "Initial",
                )
                first = commit(
                    "version: 2.1.3\nrequires:\n- Library/1.1.0\n",
                    "=Untrusted message",
                )
                commit("version: 2.1.3\n# unchanged\n", "No change")
                repo.create_head("other", initial).checkout()
                second = commit("version: 2.1.3\n", "Other branch")
                rows = find_changes(repo, "tests.yml", "2.1.2", "2.1.3")
                self.assertEqual(
                    {row["Commit Hash"] for row in rows},
                    {first.hexsha, second.hexsha},
                )
                first_row = next(
                    r for r in rows if r["Commit Hash"] == first.hexsha
                )
                self.assertEqual(
                    first_row["Requires Changes"],
                    "Removed: Library/1.0.0\nAdded: Library/1.1.0",
                )
                self.assertTrue(all(r["Files Changed"] == "tests.yml" for r in rows))
                output = root / "report.xlsx"
                write_report(rows, output)
                workbook = load_workbook(output)
                sheet = workbook["Version Changes"]
                self.assertEqual(sheet.max_row, 3)
                self.assertEqual(sheet.cell(1, 9).value, "Requires Changes")
                self.assertIn(
                    first_row["Requires Changes"],
                    [sheet.cell(i, 9).value for i in (2, 3)],
                )
                self.assertTrue(all(sheet.cell(i, 5).data_type == "s" for i in (2, 3)))
                workbook.close()

                empty = find_changes(repo, "tests.yml", "9.0.0", "9.0.1")
                write_report(empty, output)
                workbook = load_workbook(output)
                sheet = workbook["Version Changes"]
                self.assertEqual(sheet.max_row, 1)
                self.assertEqual([cell.value for cell in sheet[1]], COLUMNS)
                workbook.close()

                for content in ("other: value", "version: [", "version: banana"):
                    invalid = commit(content, "Invalid document")
                    self.assertEqual(version_at(invalid, "tests.yml"), (True, None))
                with self.assertRaises(ValueError):
                    find_changes(repo, "missing.yml", "2.1.2", "2.1.3")


if __name__ == "__main__":
    unittest.main()
