"""
Pre-commit coding-notes pattern checker for Claude Code (Windows-compatible).
Triggered via PreToolUse hook on Bash tool calls.
Reads tool input JSON from stdin, skips non-commit commands,
then checks the staged diff against known .claude/CODING_NOTES.md anti-patterns.
"""

import sys
import json
import re
import subprocess


def get_staged_diff() -> str:
    try:
        result = subprocess.run(
            ["git", "diff", "--cached"],
            capture_output=True, text=True
        )
        if result.returncode != 0:
            print(f"git diff --cached failed: {result.stderr.strip()}", file=sys.stderr)
            return ""
        return result.stdout
    except (FileNotFoundError, subprocess.SubprocessError) as e:
        print(f"Failed to run git diff: {e}", file=sys.stderr)
        return ""


def main():
    try:
        data = json.load(sys.stdin)
        command = data.get("command", "")
    except json.JSONDecodeError as e:
        print(f"pre_commit_sp_check: failed to parse hook input: {e}", file=sys.stderr)
        sys.exit(1)

    # Match git commit with optional flags/options (e.g. git -c core.editor=... commit)
    if not re.search(r"\bgit\b.*\bcommit\b", command):
        sys.exit(0)

    diff = get_staged_diff()
    if not diff:
        sys.exit(0)

    added_lines = [line for line in diff.splitlines() if line.startswith("+") and not line.startswith("+++")]

    warnings = []

    # ---------------------------------------------------------------
    # Coding-notes checks — add new entries here as CodeRabbit reviews land
    # ---------------------------------------------------------------

    # [2026-04-06] Broad except Exception
    for line in added_lines:
        if re.search(r"except\s+Exception\b", line):
            warnings.append(
                "Broad 'except Exception' detected — use specific exceptions "
                "(e.g. OSError, AttributeError). [CODING_NOTES 2026-04-06]"
            )
            break

    # [2026-04-06] Helper functions defined inside methods (8+ spaces = nested scope)
    for line in added_lines:
        if re.search(r"^\+\s{8,}def\s+_\w+", line):
            warnings.append(
                "Private helper function defined inside a method — move to class or "
                "module level. [CODING_NOTES 2026-04-06]"
            )
            break

    # [2026-04-06] Sort key puts files before dirs (isdir without not)
    for line in added_lines:
        if re.search(r"key\s*=\s*lambda.*os\.path\.isdir", line) and not re.search(r"not\s+os\.path\.isdir", line):
            warnings.append(
                "Sort key may put files before directories — "
                "use 'not os.path.isdir(...)' to sort dirs first. [CODING_NOTES 2026-04-06]"
            )
            break

    # ---------------------------------------------------------------

    if warnings:
        print()
        print(f"Coding Notes Pre-commit Check — {len(warnings)} issue(s) found:")
        for w in warnings:
            print(f"  * {w}")
        print()
        print("Review .claude/CODING_NOTES.md before proceeding. Commit is NOT blocked — "
              "fix on next commit if intentional.")
        print()

    sys.exit(0)


if __name__ == "__main__":
    main()
