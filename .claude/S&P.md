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

---

## 2026-04-10 — `modules/printing.py` (PR #4 — feat/printer-orientation)

**Review:** CodeRabbit review of orientation/auto-rotate/auto-center implementation — 2 actionable findings.
**Result:** Both fixed in `feat/printer-orientation`.

### Findings

1. **`_detect_landscape` counted all pages, ignoring print filters**
   - Scanned every page in every file; would mis-detect orientation when `print_range`, odd/even, or reverse filters meant only a subset would print
   - Fix: compute filtered indices (`_filtered_indices` helper) before calling `_detect_landscape`; pass `(entry, indices)` tuples so only pages that will actually print are counted

2. **Manual aspect-ratio scaling duplicated PyMuPDF's `keep_proportion`**
   - Manual `scale = min(cell_w/src_w, cell_h/src_h)` logic recomputed what `show_pdf_page(keep_proportion=True)` already does natively
   - Fix: pass `keep_proportion=auto_center` directly to `show_pdf_page` and use the full `cell_rect` in both cases; remove manual scaling
