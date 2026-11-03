# YAML version history report

Requires Python 3.10+ and Git installed and available on PATH.

## Windows setup (Git Bash)

Run from your repository root:

```bash
python -m venv .venv
source .venv/Scripts/activate
python -m pip install -r requirements.txt
python version_history.py --file-path test.yml --old-version 1.2.6 --new-version 1.2.7
```

For this repository's `test.yml`:

```bash
python version_history.py --file-path test.yml --old-version 2.0.1 --new-version 2.1.3
```

An absolute path with spaces also works:

```bash
python version_history.py --file-path "C:/Projects/My Repo/config/tests.yml" --old-version 2.1.2 --new-version 2.1.3
```

Run `source .venv/Scripts/activate` again when opening a new Git Bash session.

The command overwrites `version_changes_report.xlsx` in the current working
directory. Close that file in Excel before running the command again.
Its `Version Changes` worksheet has one row per dependency with these columns:

| Subsystem | Composant SW | SP01b03 | SP01b02 | evolution | Type de modification |
| --- | --- | --- | --- | --- | --- |
| Compo Alignment | asm | 4.1.1 | 4.1.0 | changed | quality change or correction without functional or architectural impact |

The row above is an illustrative example, not a claim about local history.
The third column displays component versions from `--new-version`; the fourth
displays them from `--old-version`. Their names are generated automatically
as `SP{minor:02d}b{patch:02d}`: `1.2.3` becomes `SP02b03`, and `5.15.3`
becomes `SP15b03`. The major number and prerelease/build suffix are ignored
for labels only. Numbers use at least two digits without truncation.
If the two labels coincide, both columns are retained; header comments
identify the full old/new versions. The command still accepts three inputs.

Full-line comments inside `requires` identify subsystems, for example
`# MBD Libraries`, `# Compo Alignment`, and `#Compo AngleSensor`. Entries without
a preceding section comment use `Ungrouped`. Inline comments are ignored.
Component names are extracted before `/`, and component versions before `@`.
Names must be unique within a snapshot; invalid references cause a clear error.
An absent `requires` key is treated as an empty dependency list.

Evolution is `new`, `deleted`, `changed`, or `same`. A full reference change
(including user/channel) counts as `changed`, even with identical version
numbers. Reordering entries or moving a component between subsystems alone
does not count as a reference change. The new subsystem is displayed, except
for deleted components, which retain their old subsystem. Missing versions
are blank for added/deleted components.

`Type de modification` follows the highest changed version field:

- Major: `design change : architecture`.
- Minor: `design change : functional`.
- Patch: `quality change or correction without functional or architectural impact`.
- Unchanged, added, or removed: `same`, `new`, or `deleted` respectively.

These labels are inferred from version numbers, including downgrades; they
are not a code-level impact analysis. User/channel-only changes are labeled
`reference change (user/channel)`. Suffix-only changes are labeled
`prerelease/build metadata change`. Non-SemVer versions are marked
`unclassified (non-SemVer version)` when changed.

Hover over a Subsystem cell to see the compared commit hashes, author,
date, message, changed files, and full old/new references. Multiple matching
commit pairs produce separate groups of component rows, identified by these
comments. New-snapshot order is preserved, with deleted entries appended.

Excel cells exceeding its 32,767-character limit are
truncated with a warning; unsupported control characters are removed.

## Comparison rules

- The YAML must have a top-level string `version`, such as `version: 2.1.2`.
  Versions follow SemVer: `MAJOR.MINOR.PATCH`, optionally with prerelease and
  build suffixes. Comparisons use exact strings.
- All history reachable from local branches, tags, remote-tracking refs and
  HEAD is scanned. No fetch is performed; shallow clones have limited history.
- Each commit introducing the new version is compared with its nearest
  old-version ancestor by number of parent edges. Intermediate versions are
  allowed, so `2.0.1 -> 2.1.1 -> 2.1.2 -> 2.1.3` matches a search for
  `2.0.1` to `2.1.3`. Commits inheriting the new version from any parent
  are skipped. Unrelated branches are never used as comparison baselines.
- The component table compares the two endpoint snapshots. Equally near old
  ancestors each produce a comparison, identified in the row comments.
- Commit metadata describes the commit introducing the new version.
  `Files Changed` in the cell comment lists that commit's changes against its parents,
  rather than all changes across the version range.
- Root commits and file additions have no previous version and do not match.
- The supplied path is fixed across history; renames are not followed.
  Files deleted from the working tree can still be searched in history.
- Invalid historical YAML, missing version keys, and invalid historical
  versions generate warnings and are excluded from comparisons. Missing files
  in individual commits are normal. Paths absent from all history are errors.
- A successful scan with no matches still writes a workbook with headers.
  Fatal errors print a message and return exit code 1; argument errors return 2.
