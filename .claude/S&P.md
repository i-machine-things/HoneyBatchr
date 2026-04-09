# Standards & Practices — CodeRabbit Review Log

This file records CodeRabbit recommendations so they can be applied to future changes.
Review this file before making changes to the codebase.

---

## 2026-04-09 — `.claude/hooks/pre_commit_sp_check.py` + workflow (PR #1 — chore/integrate-coderabbit-workflow)

**Review:** CodeRabbit review of initial workflow integration — 4 actionable findings.
**Result:** All 4 findings fixed in `fix/coderabbit-pr1-findings`.

### Findings

1. **Git commit regex too narrow**
   - `re.search(r"git\s+commit", command)` misses `git -c flag=val commit`
   - Fix: use `r"\bgit\b.*\bcommit\b"` to match any git invocation with commit subcommand

2. **Broad `except Exception` in hook**
   - `get_staged_diff` caught all exceptions silently; JSON parse caught all exceptions
   - Fix: use `except (FileNotFoundError, subprocess.SubprocessError)` for subprocess; `except json.JSONDecodeError` for JSON; check `result.returncode` and log errors to stderr

3. **`settings.local.json` tracked with machine-specific wildcard paths**
   - File contained absolute paths and broad wildcards not portable across machines
   - Fix: gitignore `settings.local.json`; add `settings.local.json.example` with minimal portable entries

4. **PyInstaller spec file gitignored — CI workflow broken**
   - `build/batch_print.spec` is inside the gitignored `build/` directory; workflow would fail
   - Fix: copy spec to `build_scripts/HoneyBatchr.spec` (tracked); update workflow to reference `build_scripts/HoneyBatchr.spec`
