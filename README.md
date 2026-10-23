# Check Dependency Tool

## Overview
`checkdepend.py` looks up upstream Linux commits that fix a given commit, and it can also check whether a local backport is missing dependency commits from a Torvalds tree.

## Find Fixes patches
Run the tool with no arguments:

```bash
./checkdepend.py
```

It asks for:

1. The path to a Linux upstream tree.
2. Upstream commit IDs, one per line. Type `done` when the list is finished.

It then switches that tree to `master`, pulls the latest code, and lists commits whose message carries a `Fixes:` tag for each ID.

```bash
$ ./checkdepend.py
Enter path to Linux upstream source: /home/user/linux
Paste upstream commit IDs (one per line). Type 'done' when finished:
abc1234
def5678
done
```

A commit with no matching `Fixes:` tag is reported as `No Fixes patches found`.

## Backport dependency check
`./checkdepend.py --backport` is the previous check. It compares commits in your kernel tree with a Torvalds tree and reports whether a later commit that mentions yours is already present.

```bash
$ ./checkdepend.py --backport
Enter path to your kernel source: /home/user/my-kernel
Enter path to Torvalds kernel source: /home/user/linux-stable
Enter commits (one per line). Type 'done' when finished:
abc1234
def5678
done
```

### Output
```bash
Commit: abc123456789ab - Fix memory leak in driver
  PASS -> No dependencies found

Commit: def56781234567 - Add new feature to subsystem
  FAIL -> new bugfix needed
    * 1234567890abcd Dependency commit subject
```

The backport check writes `.stable_log`, `.user_log`, `.full_commits`, and `.dep_log` in the current directory.

- **PASS** means no dependencies were found, or every dependency subject is already in your tree.
- **FAIL** means a dependency commit still needs to be backported.

## Excel sheet
After either check finishes, the tool asks:

```text
Generate excel sheet? [y/N]:
```

Press Enter to skip. `y` or `yes` writes `checkdepend-fixes-<date>.xlsx` or `checkdepend-backport-<date>.xlsx` in the current directory. This needs `openpyxl`.

## Install
`--install` compiles a binary with `gcc` and installs it as `checkdepend`. It uses `/usr/local/bin` when that directory is writable, and `~/.local/bin` otherwise.

```bash
./checkdepend.py --install
./checkdepend.py --uninstall
./checkdepend.py --reinstall
```

`--reinstall` compiles a new binary and replaces the installed one. `gcc` and `python3` are required on the machine that runs the install.
