# YAML version history report

Requires Python 3.10+ and Git installed and available on PATH.

## Windows setup (Git Bash)

Run from your repository root:

```bash
python -m venv .venv
source .venv/Scripts/activate
python -m pip install -r requirements.txt
python version_history.py --file-path config/tests.yml --old-version 2.1.2 --new-version 2.1.3
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
Its `Version Changes` worksheet includes commit hash, author name and email,
committer date (ISO 8601 with timezone), full commit message, old/new versions,
and changed paths. Excel cells exceeding its 32,767-character limit are
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
- `Requires Changes` shows added/removed entries between the two endpoint
  snapshots. `Old Version Commit` identifies the baseline. Equally near old
  ancestors are all included and labeled in the same row for merge histories.
- Commit metadata describes the commit introducing the new version.
  `Files Changed` still lists that commit's changes against its parents,
  rather than all changes across the version range.
- Root commits and file additions have no previous version and do not match.
- The supplied path is fixed across history; renames are not followed.
  Files deleted from the working tree can still be searched in history.
- Invalid historical YAML, missing version keys, and invalid historical
  versions generate warnings and are excluded from comparisons. Missing files
  in individual commits are normal. Paths absent from all history are errors.
- A successful scan with no matches still writes a workbook with headers.
  Fatal errors print a message and return exit code 1; argument errors return 2.
