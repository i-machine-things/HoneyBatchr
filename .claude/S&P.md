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

---

## 2026-04-12 — `modules/printing.py` + `modules/app.py` (PR #5 — feat/paper-size-page-setting)

**Review:** CodeRabbit review of Page Setting dialog implementation — 3 actionable findings.
**Result:** All resolved in `feat/paper-size-page-setting`.

### Findings

1. **Negative page margins accepted silently**
   - `page_margin_*` values were converted to points without a sign check; negative inputs produced incorrect (expanded) printable area rather than an error
   - Fix: validate all four margin params are `>= 0` before conversion; raise `ValueError` naming the offending param and value

2. **Page Setting dialog orientation/paper_size not wired end-to-end**
   - `compose_nup_pdf()` was still reading `self.orientation_combo.currentText()` instead of the stored `page_setting_orientation` config key; `paper_size` was not propagated to `print_pdf_qt()` / `QPrinter`
   - Fix (commit 45d1016): use `self.config.get("page_setting_orientation", ...)` at the call site; pass `paper_size` to `print_pdf_qt()`; apply `QPageSize` on `QPrinter` before `painter.begin()`

3. **Combined margins not validated against paper dimensions in dialog**
   - Individual spinboxes bounded to 4 in. each, but `left+right` or `top+bottom` could still exceed the selected sheet side (e.g., A5 at 4"+4"), yielding a non-positive printable area during composition
   - Fix: override `PageSettingDialog.accept()`; extract `_selected_sheet_size()` helper; compare combined margins to sheet dimensions and show a `QMessageBox.warning` instead of closing if invalid
