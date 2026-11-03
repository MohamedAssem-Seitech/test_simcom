"""Integration checks using temporary repositories and workbooks."""

import tempfile
import unittest
from pathlib import Path

from git import Actor, Repo
from openpyxl import load_workbook

from version_history import (
    COLUMNS, compare_components, find_changes, parse_components, version_at,
    write_report, modification_type,
)


class HistoryTests(unittest.TestCase):
    """Exercise real Git snapshots and the generated Excel report."""

    def test_modification_types(self) -> None:
        """Classify version fields by precedence without inventing impact."""
        cases = [
            ("7.1.0", "8.1.0", "changed", "design change : architecture"),
            ("2.0.0", "2.1.0", "changed", "design change : functional"),
            ("1.0.0", "1.0.1", "changed",
             "quality change or correction without functional or architectural impact"),
            ("1.0.0", "1.0.0", "same", "same"),
            ("", "1.0.0", "new", "new"),
            ("1.0.0", "", "deleted", "deleted"),
            ("1.0.0", "2.1.1", "changed", "design change : architecture"),
            ("2.1.1", "2.0.0", "changed", "design change : functional"),
            ("1.0.0", "1.0.0", "changed", "reference change (user/channel)"),
            ("1.0.0-rc.1", "1.0.0", "changed",
             "prerelease/build metadata change"),
            ("1", "2", "changed", "unclassified (non-SemVer version)"),
        ]
        for old, new, evolution, expected in cases:
            with self.subTest(old=old, new=new, evolution=evolution):
                self.assertEqual(
                    modification_type(old, new, evolution), expected,
                )

    def test_component_comparison(self) -> None:
        """Preserve comment groups and distinguish all four evolution states."""
        old = '''name: L060_HWA
version: 2.1.2
requires:
  # Compo Alignment
  - "asm/4.1.0@jenkins/l060"
  - "alt/1.1.0@jenkins/l060"
  #Compo AngleSensor
  - "das/4.0.0@jenkins/l060"
  - "cas/7.0.1@jenkins/old"
'''
        new = '''name: L060_HWA
version: 2.1.3
requires:
  # Compo Alignment
  - "asm/4.1.1@jenkins/l060"  # ordinary inline comment
  - "alt/1.1.0@jenkins/l060"
  #Compo AngleSensor
  - "cac_hwa/4.0.0@jenkins/l060"
  - "cas/7.0.1@jenkins/new"
'''
        rows = {r["Composant SW"]: r for r in compare_components(old, new)}
        self.assertEqual(rows["asm"]["Subsystem"], "Compo Alignment")
        self.assertEqual(rows["asm"]["SP01b03"], "4.1.1")
        self.assertEqual(rows["asm"]["SP01b02"], "4.1.0")
        self.assertEqual(rows["asm"]["evolution"], "changed")
        self.assertEqual(rows["alt"]["evolution"], "same")
        self.assertEqual(rows["cac_hwa"]["evolution"], "new")
        self.assertEqual(rows["cac_hwa"]["SP01b02"], "")
        self.assertEqual(rows["das"]["evolution"], "deleted")
        self.assertEqual(rows["das"]["SP01b03"], "")
        self.assertEqual(rows["das"]["Subsystem"], "Compo AngleSensor")
        self.assertEqual(rows["cas"]["evolution"], "changed")
        self.assertEqual(parse_components("requires: []"), {})
        with self.assertRaises(ValueError):
            parse_components("requires: [asm/1.0.0, asm/2.0.0]")
        with self.assertRaises(ValueError):
            parse_components("requires: [invalid]")

    def test_intermediate_versions(self) -> None:
        """Compare endpoints across intermediate versions on actual ancestry."""
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            with Repo.init(root) as repo:
                actor = Actor("Test", "test@example.com")
                commits = []
                for version, dependency in (
                    ("2.0.1", "Library/1"),
                    ("2.1.1", "Library/2"),
                    ("2.1.2", "Library/3"),
                    ("2.1.3", "Library/4"),
                    ("2.1.3", "Library/4"),
                ):
                    (root / "tests.yml").write_text(
                        f"version: {version}\nrequires:\n- {dependency}\n",
                        encoding="utf-8",
                    )
                    repo.index.add(["tests.yml"])
                    commits.append(repo.index.commit(
                        version, author=actor, committer=actor,
                    ))
                rows = find_changes(repo, "tests.yml", "2.0.1", "2.1.3")
                self.assertEqual(len(rows), 1)
                self.assertEqual(rows[0]["Commit Hash"], commits[3].hexsha)
                self.assertEqual(
                    rows[0]["Old Version Commit"], commits[0].hexsha,
                )
                self.assertEqual(
                    rows[0]["Requires Changes"],
                    "Removed: Library/1\nAdded: Library/4",
                )
                # A version on a descendant cannot serve as an old ancestor.
                self.assertEqual(
                    find_changes(repo, "tests.yml", "2.1.3", "2.0.1"), [],
                )

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
                self.assertEqual([c.value for c in sheet[1]], COLUMNS)
                self.assertEqual(sheet.max_column, 6)
                self.assertIn(
                    "design change : functional",
                    [sheet.cell(i, 6).value for i in (2, 3)],
                )
                self.assertEqual(
                    {sheet.cell(i, 5).value for i in (2, 3)},
                    {"changed", "deleted"},
                )
                for i in (2, 3):
                    self.assertEqual(sheet.cell(i, 2).value, "Library")
                    self.assertEqual(sheet.cell(i, 2).data_type, "s")
                    self.assertIn(
                        "Commit Hash:", sheet.cell(i, 1).comment.text,
                    )
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
