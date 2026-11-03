# YAML version history report

Requires Python 3.10+ and Git installed and available on PATH.

## Windows setup (PowerShell or Command Prompt)

Run from your repository root:

```powershell
py -3 -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe version_history.py --file-path config/tests.yml --old-version 2.1.2 --new-version 2.1.3
```

For this repository's `test.yml`:

```powershell
.venv\Scripts\python.exe version_history.py --file-path test.yml --old-version 2.1.2 --new-version 2.1.3
```

An absolute path with spaces also works:

```powershell
.venv\Scripts\python.exe version_history.py --file-path "C:\Projects\My Repo\config\tests.yml" --old-version 2.1.2 --new-version 2.1.3
```

No virtual environment activation is needed. If `py` is unavailable, use your
installed Python 3.10+ executable instead.

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
- Every commit is compared with its actual parents, not its neighbor in a
  date-sorted log. Merges match if any parent has the requested old version,
  and appear once. Changed files are the union of diffs against all parents.
- Root commits and file additions have no previous version and do not match.
- The supplied path is fixed across history; renames are not followed.
  Files deleted from the working tree can still be searched in history.
- Invalid historical YAML, missing version keys, and invalid historical
  versions generate warnings and are excluded from comparisons. Missing files
  in individual commits are normal. Paths absent from all history are errors.
- A successful scan with no matches still writes a workbook with headers.
  Fatal errors print a message and return exit code 1; argument errors return 2.
