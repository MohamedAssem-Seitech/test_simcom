"""Compare YAML versions across Git history and export to Excel.

Requires Python 3.10+, Git, and the packages in requirements.txt.
Relative file paths are resolved against the current working directory.
All commits reachable from local refs (including remote-tracking refs) and
HEAD are scanned. Compare commits introducing the new version with their
nearest old-version ancestors, allowing intermediate versions.
Alternatively, supply two commit IDs to compare their snapshots directly,
including commits on different branches. No checkout or fetch is performed.
"""

from __future__ import annotations

import argparse
import difflib
import re
import sys
from pathlib import Path
from typing import Any

import pandas as pd
import yaml
from git import Commit, Repo
from git.exc import GitError
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill


VERSION_PATTERN = re.compile(
    r"(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)"
    r"(?:-((?:0|[1-9]\d*|\d*[A-Za-z-][0-9A-Za-z-]*)"
    r"(?:\.(?:0|[1-9]\d*|\d*[A-Za-z-][0-9A-Za-z-]*))*))?"
    r"(?:\+([0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*))?",
    flags=re.ASCII,
)
NEW_RELEASE_LABEL = "New Component Version"
OLD_RELEASE_LABEL = "Old Component Version"
COLUMNS = [
    "Subsystem", "Composant SW", NEW_RELEASE_LABEL, OLD_RELEASE_LABEL,
    "evolution", "Type de modification",
]


def parse_components(source: str) -> dict[str, dict[str, str]]:
    """Read dependency references and preceding full-line section comments.

    SafeLoader validates values; YAML node line numbers associate comments
    with entries without mistaking inline comments for subsystem headings.
    Component names must be unique within one snapshot.
    """
    document = yaml.safe_load(source)
    if not isinstance(document, dict):
        raise ValueError("Expected a YAML mapping")
    if "requires" not in document:
        return {}
    if not isinstance(document["requires"], list):
        raise ValueError("'requires' must be a list of dependency strings")
    root = yaml.compose(source, Loader=yaml.SafeLoader)
    assert isinstance(root, yaml.MappingNode)
    key, sequence = next(
        (key, value) for key, value in root.value if key.value == "requires"
    )
    if not isinstance(sequence, yaml.SequenceNode):
        raise ValueError("'requires' must be a YAML sequence")
    lines = source.splitlines()
    cursor = key.end_mark.line + 1
    subsystem = "Ungrouped"
    components = {}
    for node, reference in zip(sequence.value, document["requires"]):
        for line in lines[cursor:node.start_mark.line]:
            stripped = line.strip()
            if stripped.startswith("#") and stripped[1:].strip():
                subsystem = stripped[1:].strip()
        cursor = node.end_mark.line + 1
        if not isinstance(reference, str):
            raise ValueError("Each requires entry must be a string")
        match = re.fullmatch(r"([^/\s]+)/([^@\s]+)(?:@([^\s]+))?", reference)
        if not match:
            raise ValueError(f"Invalid dependency reference: {reference!r}")
        name, version, _ = match.groups()
        if name in components:
            raise ValueError(f"Duplicate component name in requires: {name}")
        components[name] = {
            "subsystem": subsystem, "version": version, "reference": reference,
        }
    return components


def modification_type(old_version: str, new_version: str, evolution: str) -> str:
    """Classify the highest changed version field, including downgrades.

    Reference-only and suffix-only changes have their own labels because
    their architectural or functional impact cannot be inferred.
    """
    if evolution != "changed":
        return evolution
    if old_version == new_version:
        return "reference change (user/channel)"
    old_match = VERSION_PATTERN.fullmatch(old_version)
    new_match = VERSION_PATTERN.fullmatch(new_version)
    if old_match is None or new_match is None:
        return "unclassified (non-SemVer version)"
    labels = (
        "design change : architecture",
        "design change : functional",
        "quality change or correction without functional or architectural impact",
    )
    for index, label in enumerate(labels, start=1):
        if old_match.group(index) != new_match.group(index):
            return label
    return "prerelease/build metadata change"


def compare_components(old_source: str, new_source: str) -> list[dict[str, str]]:
    """Build one row per dependency, keeping new order then deleted entries."""
    old = parse_components(old_source)
    new = parse_components(new_source)
    rows = []
    for name in dict.fromkeys([*new, *old]):
        before, after = old.get(name), new.get(name)
        if before is None:
            evolution = "new"
        elif after is None:
            evolution = "deleted"
        elif before["reference"] != after["reference"]:
            evolution = "changed"
        else:
            evolution = "same"
        component = after if after is not None else before
        assert component is not None
        rows.append({
            "Subsystem": component["subsystem"],
            "Composant SW": name,
            NEW_RELEASE_LABEL: after["version"] if after else "",
            OLD_RELEASE_LABEL: before["version"] if before else "",
            "evolution": evolution,
            "Type de modification": modification_type(
                before["version"] if before else "",
                after["version"] if after else "", evolution,
            ),
            "_old_reference": before["reference"] if before else "",
            "_new_reference": after["reference"] if after else "",
        })
    return rows


def valid_version(value: str) -> str:
    """Validate a SemVer string, preserving its exact spelling."""
    if not VERSION_PATTERN.fullmatch(value):
        raise argparse.ArgumentTypeError(
            f"Invalid version {value!r}; expected MAJOR.MINOR.PATCH, "
            "optionally with a prerelease or build suffix (e.g. 2.1.3-rc.1)."
        )
    return value


def valid_selector(value: str) -> str:
    """Accept a semantic version or a short/full hexadecimal commit ID."""
    if VERSION_PATTERN.fullmatch(value) or re.fullmatch(r"[0-9a-fA-F]{4,40}", value):
        return value
    raise argparse.ArgumentTypeError(
        "Expected a version such as 2.1.3 or a commit ID such as c2f9afb2."
    )


def parse_arguments() -> argparse.Namespace:
    """Parse the three required user inputs."""
    parser = argparse.ArgumentParser(description=__doc__, allow_abbrev=False)
    parser.add_argument("--file-path", required=True, type=Path)
    parser.add_argument("--old-version", required=True, type=valid_selector)
    parser.add_argument("--new-version", required=True, type=valid_selector)
    args = parser.parse_args()
    if args.old_version == args.new_version:
        parser.error("--old-version and --new-version must differ.")
    if bool(VERSION_PATTERN.fullmatch(args.old_version)) != bool(
        VERSION_PATTERN.fullmatch(args.new_version)
    ):
        parser.error("Use two version numbers or two commit IDs, not a mixture.")
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


def requires_changes(parent: Commit, commit: Commit, file_path: str) -> str:
    """Describe removed/added requires entries from validated YAML snapshots."""
    def lines_at(revision: Commit) -> list[str]:
        blob = revision.tree / file_path
        document = yaml.safe_load(blob.data_stream.read().decode("utf-8-sig"))
        if "requires" not in document:
            return ["[requires key absent]"]
        requires = document["requires"]
        if isinstance(requires, list) and all(
            isinstance(item, str) for item in requires
        ):
            return requires
        return yaml.safe_dump(requires, sort_keys=True).splitlines()

    changes = []
    for line in difflib.ndiff(lines_at(parent), lines_at(commit)):
        if line.startswith("- "):
            changes.append(f"Removed: {line[2:]}")
        elif line.startswith("+ "):
            changes.append(f"Added: {line[2:]}")
    return "\n".join(changes) or "No changes"


def nearest_old_commits(
    commit: Commit, versions: dict[str, str | None], old_version: str,
) -> list[Commit]:
    """Find old-version ancestors at the smallest parent-edge distance.

    Breadth-first traversal follows actual ancestry, never unrelated branches.
    Equally close ancestors are retained so merge comparisons are explicit.
    """
    frontier = list(commit.parents)
    seen = {commit.hexsha}
    while frontier:
        level = []
        for ancestor in frontier:
            if ancestor.hexsha not in seen:
                seen.add(ancestor.hexsha)
                level.append(ancestor)
        matches = [
            ancestor for ancestor in level
            if versions.get(ancestor.hexsha) == old_version
        ]
        if matches:
            return matches
        frontier = [parent for ancestor in level for parent in ancestor.parents]
    return []


def find_changes(
    repo: Repo, file_path: str, old_version: str, new_version: str,
) -> list[dict[str, Any]]:
    """Compare new-version introductions with nearest old-version ancestors."""
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
        # Do not repeat a version already present in a parent snapshot.
        if any(versions.get(p.hexsha) == new_version for p in commit.parents):
            continue
        baselines = nearest_old_commits(commit, versions, old_version)
        if not baselines:
            continue
        requires_diffs = []
        component_rows = []
        for baseline in baselines:
            difference = requires_changes(baseline, commit, file_path)
            if len(baselines) > 1:
                difference = f"Old commit {baseline.hexsha}:\n{difference}"
            requires_diffs.append(difference)
            old_source = (baseline.tree / file_path).data_stream.read().decode(
                "utf-8-sig",
            )
            new_source = (commit.tree / file_path).data_stream.read().decode(
                "utf-8-sig",
            )
            for component_row in compare_components(old_source, new_source):
                component_row["_baseline"] = baseline.hexsha
                component_rows.append(component_row)
        matches.append({
            "Commit Hash": commit.hexsha,
            "Author Name": commit.author.name or "",
            "Author Email": commit.author.email or "",
            "Commit Date": commit.committed_datetime.isoformat(),
            "Commit Message": commit.message,
            "Old Version": old_version,
            "New Version": new_version,
            "Files Changed": changed_files(commit),
            "Requires Changes": "\n\n".join(requires_diffs),
            "Old Version Commit": "\n".join(c.hexsha for c in baselines),
            "Components": component_rows,
        })
    return matches


def compare_commits(
    repo: Repo, file_path: str, old_id: str, new_id: str,
) -> list[dict[str, Any]]:
    """Compare explicitly selected commits without any ancestry requirement."""
    snapshots = []
    for identifier in (old_id, new_id):
        if not re.fullmatch(r"[0-9a-fA-F]{4,40}", identifier):
            raise ValueError(f"Invalid commit ID: {identifier!r}")
        try:
            sha = repo.git.rev_parse("--verify", f"{identifier.lower()}^{{commit}}")
        except GitError as exc:
            raise ValueError(
                f"Cannot resolve commit {identifier!r}. Use a unique commit ID "
                "available in this local repository; fetch its branch if needed."
            ) from exc
        commit = repo.commit(sha)
        present, version = version_at(commit, file_path)
        if not present:
            raise ValueError(f"{file_path!r} is missing at commit {identifier}.")
        if version is None:
            raise ValueError(
                f"{file_path!r} at {identifier} must contain valid YAML with a "
                "top-level SemVer version to generate the SP column header."
            )
        source = (commit.tree / file_path).data_stream.read().decode("utf-8-sig")
        snapshots.append((commit, version, source))
    (old, old_version, old_source), (new, new_version, new_source) = snapshots
    if old.hexsha == new.hexsha:
        raise ValueError("Both IDs resolve to the same commit; select two commits.")
    components = compare_components(old_source, new_source)
    for component in components:
        component["_baseline"] = old.hexsha
    paths = {
        path for change in old.diff(new)
        for path in (change.a_path, change.b_path) if path
    }
    return [{
        "Commit Hash": new.hexsha,
        "Author Name": new.author.name or "",
        "Author Email": new.author.email or "",
        "Commit Date": new.committed_datetime.isoformat(),
        "Commit Message": new.message,
        "Old Version": old_version,
        "New Version": new_version,
        "Files Changed": "\n".join(sorted(paths)),
        "Old Version Commit": old.hexsha,
        "Components": components,
    }]


def release_label(version: str) -> str:
    """Format a release's minor and patch fields as SPNNbNN."""
    valid_version(version)
    match = VERSION_PATTERN.fullmatch(version)
    assert match is not None
    return f"SP{int(match.group(2)):02d}b{int(match.group(3)):02d}"


def report_headers(old_version: str, new_version: str) -> list[str]:
    """Return display headers, retaining distinct internal keys for columns."""
    return [
        "Subsystem", "Composant SW", release_label(new_version),
        release_label(old_version), "evolution", "Type de modification",
    ]


def write_report(
    rows: list[dict[str, Any]], output: Path,
    old_version: str, new_version: str,
) -> None:
    """Write component columns, with source metadata in cell comments."""
    cleaned = []
    comments = []
    for match in rows:
        for row in match["Components"]:
            cleaned_row = {}
            for key in COLUMNS:
                text = ILLEGAL_CHARACTERS_RE.sub("", str(row[key]))
                if len(text) > 32767:
                    print(f"Warning: truncating {key} to Excel's cell limit.",
                          file=sys.stderr)
                    text = text[:32750] + "\n[truncated]"
                cleaned_row[key] = text
            cleaned.append(cleaned_row)
            metadata = "\n".join(
                f"{key}: {value}" for key, value in match.items()
                if key not in {"Components", "Requires Changes"}
            )
            comments.append(ILLEGAL_CHARACTERS_RE.sub(
                "", f"{metadata}\nCompared old commit: {row['_baseline']}\n"
                f"Old reference: {row['_old_reference']}\n"
                f"New reference: {row['_new_reference']}",
            ))
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pd.DataFrame(cleaned, columns=COLUMNS).to_excel(
            writer, sheet_name="Version Changes", index=False,
            header=report_headers(old_version, new_version),
        )
        sheet = writer.sheets["Version Changes"]
        # Commit messages and filenames are untrusted text, never formulas.
        for row in sheet.iter_rows():
            for cell in row:
                cell.data_type = "s"
                cell.alignment = Alignment(wrap_text=True, vertical="top")
        for cell in sheet[1]:
            cell.fill = PatternFill("solid", fgColor="17365D")
            cell.font = Font(color="FFFFFF", bold=True)
        colors = {
            "changed": "FFF2CC", "new": "E2F0D9",
            "deleted": "FCE4D6", "same": "E7E6E6",
        }
        for index, comment in enumerate(comments, start=2):
            sheet.cell(index, 1).comment = Comment(comment, "Version history")
            status = sheet.cell(index, 5)
            status.fill = PatternFill("solid", fgColor=colors[status.value])
        sheet.cell(1, 3).comment = Comment(
            f"Component version in --new-version {new_version}.",
            "Version history",
        )
        sheet.cell(1, 4).comment = Comment(
            f"Component version in --old-version {old_version}.",
            "Version history",
        )
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for column, width in zip(
            "ABCDEF", (34, 34, 24, 24, 18, 75),
        ):
            sheet.column_dimensions[column].width = width


def main() -> int:
    """Run the scan and report actionable errors without a traceback."""
    args = parse_arguments()
    commit_mode = VERSION_PATTERN.fullmatch(args.old_version) is None
    output = Path.cwd() / "version_changes_report.xlsx"
    try:
        repo, relative_path = locate_repository(args.file_path)
        with repo:
            if commit_mode:
                rows = compare_commits(
                    repo, relative_path, args.old_version, args.new_version,
                )
                old_version = rows[0]["Old Version"]
                new_version = rows[0]["New Version"]
            else:
                rows = find_changes(
                    repo, relative_path, args.old_version, args.new_version,
                )
                old_version, new_version = args.old_version, args.new_version
        write_report(rows, output, old_version, new_version)
    except (GitError, OSError, ValueError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    if not rows:
        print(f"No commits found where version changed from "
              f"{args.old_version} to {args.new_version}.")
    if commit_mode:
        print(f"Compared commits: {rows[0]['Old Version Commit']} -> "
              f"{rows[0]['Commit Hash']}")
    else:
        print(f"Matching commits: {len(rows)}")
    print(f"Component rows: {sum(len(row['Components']) for row in rows)}")
    print(f"Excel report: {output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
