"""Find exact YAML version transitions in Git history and export to Excel.

Requires Python 3.10+, Git, and the packages in requirements.txt.
Relative file paths are resolved against the current working directory.
All commits reachable from local refs (including remote-tracking refs) and
HEAD are scanned. A merge matches if any parent has the requested old version;
each matching commit is reported once. Root commits have no previous version.
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from typing import Any

import pandas as pd
import yaml
from git import Commit, Repo
from git.exc import GitError
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE


VERSION_PATTERN = re.compile(
    r"(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)"
    r"(?:-((?:0|[1-9]\d*|\d*[A-Za-z-][0-9A-Za-z-]*)"
    r"(?:\.(?:0|[1-9]\d*|\d*[A-Za-z-][0-9A-Za-z-]*))*))?"
    r"(?:\+([0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*))?",
    flags=re.ASCII,
)
COLUMNS = [
    "Commit Hash", "Author Name", "Author Email", "Commit Date",
    "Commit Message", "Old Version", "New Version", "Files Changed",
]


def valid_version(value: str) -> str:
    """Validate a SemVer string, preserving its exact spelling."""
    if not VERSION_PATTERN.fullmatch(value):
        raise argparse.ArgumentTypeError(
            f"Invalid version {value!r}; expected MAJOR.MINOR.PATCH, "
            "optionally with a prerelease or build suffix (e.g. 2.1.3-rc.1)."
        )
    return value


def parse_arguments() -> argparse.Namespace:
    """Parse the three required user inputs."""
    parser = argparse.ArgumentParser(description=__doc__, allow_abbrev=False)
    parser.add_argument("--file-path", required=True, type=Path)
    parser.add_argument("--old-version", required=True, type=valid_version)
    parser.add_argument("--new-version", required=True, type=valid_version)
    args = parser.parse_args()
    if args.old_version == args.new_version:
        parser.error("--old-version and --new-version must differ.")
    return args


def locate_repository(file_path: Path) -> tuple[Repo, str]:
    """Find the containing working repository and Git-relative file path."""
    path = file_path.expanduser().resolve()
    # Locate from the closest existing ancestor to support deleted files.
    ancestor = path.parent
    while not ancestor.exists() and ancestor != ancestor.parent:
        ancestor = ancestor.parent
    repo = Repo(ancestor, search_parent_directories=True)
    if repo.bare or repo.working_tree_dir is None:
        repo.close()
        raise ValueError("A non-bare Git working repository is required.")
    try:
        relative = path.relative_to(Path(repo.working_tree_dir).resolve())
        if path.exists() and not path.is_file():
            raise ValueError(f"Not a file: {path}")
        if relative.parts and relative.parts[0] == ".git":
            raise ValueError("The YAML file must be outside the .git directory.")
    except ValueError:
        repo.close()
        raise
    return repo, relative.as_posix()


def version_at(commit: Commit, file_path: str) -> tuple[bool, str | None]:
    """Read a committed blob safely; return presence and a valid version.

    Absent files are normal in history. Invalid documents are warned about
    and cannot serve as either endpoint of a matching transition.
    """
    try:
        blob = commit.tree / file_path
    except KeyError:
        return False, None
    try:
        if blob.type != "blob":
            raise ValueError("path is not a file")
        document = yaml.safe_load(blob.data_stream.read().decode("utf-8-sig"))
        if not isinstance(document, dict) or "version" not in document:
            raise ValueError("missing top-level 'version' key")
        value = document["version"]
        if not isinstance(value, str):
            raise ValueError("'version' must be a SemVer string")
        return True, valid_version(value)
    except (UnicodeError, yaml.YAMLError, ValueError,
            argparse.ArgumentTypeError) as exc:
        print(f"Warning: {commit.hexsha[:12]}: {exc}", file=sys.stderr)
        return True, None


def changed_files(commit: Commit) -> str:
    """Return the union of changed paths against all parents of a commit."""
    paths: set[str] = set()
    for parent in commit.parents:
        for change in parent.diff(commit):
            # Include both source and destination for renamed paths.
            paths.update(p for p in (change.a_path, change.b_path) if p)
    return "\n".join(sorted(paths))


def find_changes(
    repo: Repo, file_path: str, old_version: str, new_version: str,
) -> list[dict[str, Any]]:
    """Compare parent-child snapshots across all locally available history."""
    if repo.git.rev_parse("--is-shallow-repository").strip() == "true":
        print("Warning: shallow clone; only available history is scanned.",
              file=sys.stderr)
    revisions = ["--all"]
    if repo.head.is_valid():
        revisions.append("HEAD")
    commits = list(repo.iter_commits(rev=revisions))
    versions: dict[str, str | None] = {}
    found_file = False
    for commit in commits:
        present, version = version_at(commit, file_path)
        found_file = found_file or present
        versions[commit.hexsha] = version
    if not found_file:
        raise ValueError(
            f"File {file_path!r} was not found in available Git history. "
            "Check the path and ensure the file has been committed."
        )
    matches = []
    for commit in commits:
        if versions[commit.hexsha] != new_version:
            continue
        if not any(versions.get(p.hexsha) == old_version for p in commit.parents):
            continue
        matches.append({
            "Commit Hash": commit.hexsha,
            "Author Name": commit.author.name or "",
            "Author Email": commit.author.email or "",
            "Commit Date": commit.committed_datetime.isoformat(),
            "Commit Message": commit.message,
            "Old Version": old_version,
            "New Version": new_version,
            "Files Changed": changed_files(commit),
        })
    return matches


def write_report(rows: list[dict[str, Any]], output: Path) -> None:
    """Write text cells safely, retaining headers even for an empty report."""
    cleaned = []
    for row in rows:
        cleaned_row = {}
        for key, value in row.items():
            text = ILLEGAL_CHARACTERS_RE.sub("", str(value))
            if len(text) > 32767:
                print(f"Warning: truncating {key} to Excel's cell limit.",
                      file=sys.stderr)
                text = text[:32750] + "\n[truncated]"
            cleaned_row[key] = text
        cleaned.append(cleaned_row)
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pd.DataFrame(cleaned, columns=COLUMNS).to_excel(
            writer, sheet_name="Version Changes", index=False,
        )
        sheet = writer.sheets["Version Changes"]
        # Commit messages and filenames are untrusted text, never formulas.
        for row in sheet.iter_rows():
            for cell in row:
                cell.data_type = "s"
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for column, width in zip("ABCDEFGH", (44, 24, 32, 28, 70, 20, 20, 60)):
            sheet.column_dimensions[column].width = width


def main() -> int:
    """Run the scan and report actionable errors without a traceback."""
    args = parse_arguments()
    output = Path.cwd() / "version_changes_report.xlsx"
    try:
        repo, relative_path = locate_repository(args.file_path)
        with repo:
            rows = find_changes(
                repo, relative_path, args.old_version, args.new_version,
            )
        write_report(rows, output)
    except (GitError, OSError, ValueError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    if not rows:
        print(f"No commits found where version changed from "
              f"{args.old_version} to {args.new_version}.")
    print(f"Matching commits: {len(rows)}")
    print(f"Excel report: {output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
