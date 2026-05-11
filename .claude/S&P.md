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

---

## 2026-04-12 — `modules/printing.py` + `modules/app.py` (PR #6 — feat/manual-duplex)

**Review:** CodeRabbit review of manual duplex two-pass implementation — 3 temp file leak findings.
**Result:** All 3 fixed.

### Findings

1. **`tmp` leaks if `split_for_manual_duplex()` raises before `try` block**
   - `front_path`/`back_path` uninitialized when `split_for_manual_duplex` was called outside `try/finally`; if it raised, `tmp` was never unlinked
   - Fix: initialize `front_path = None`, `back_path = None`; move `split_for_manual_duplex` call inside the `try` block so `finally` always runs

2. **`front_path` leaks if `_subset(back_indices)` raises**
   - `_subset(front_indices)` succeeded but `_subset(back_indices)` could raise; front temp file left behind
   - Fix: wrap back-side `_subset` call in `try/except`; unlink `front_path` before re-raising

3. **Temp file leaks inside `_subset` if `sub.save()` fails**
   - On disk-full or I/O error, the temp file was created but never cleaned up
   - Fix: wrap `sub.save()` in `try/except`; close `sub` and unlink the temp file before re-raising; mirrors the pattern in `compose_nup_pdf`

4. **Nitpick: `tuple` return annotation too generic on `split_for_manual_duplex`**
   - Generic `tuple` gives IDEs and type checkers no useful information
   - Fix: annotate as `tuple[str, str | None]` — no import needed (Python 3.10+ built-in generics)

5. **`_toggle_duplex()` called before `manual_duplex_check` exists**
   - Init called `_toggle_duplex(...)` at line 625, but `manual_duplex_check` wasn't created until line 627; if duplex starts enabled this raises `AttributeError`
   - Fix: move `_toggle_duplex(...)` call to after both `manual_duplex_check` and `reverse_back_check` are created

6. **`tmp` never unlinked in non-manual print path**
   - The hardware-duplex `else` branch printed `tmp` but never deleted it, leaking one temp PDF per print run
   - Fix: wrap `print_pdf_qt` in `try/finally`; unlink `tmp` unconditionally

7. **Manual duplex silently produced mixed one-pass/two-pass jobs**
   - `manual_duplex_check` only affected `fitz_entries`; `other_entries` (DOCX, XLS, etc.) still went through ShellExecute/`lp` one-pass, leaving the back side unprinted for those files
   - Fix: raise `ValueError` early if `other_entries` is non-empty when manual duplex is enabled; surfaces as a print error before any pages are sent

8. **Cancel on reload dialog reported as full success**
   - User clicking Cancel on the back-side reload prompt caused status bar to show "Sent N file(s) to printer" as if the job completed
   - Fix: track `manual_duplex_canceled`; show "front sides sent — back-side pass canceled" instead of the normal success message

---

## 2026-04-12 — `.github/workflows/build-release.yml` + `README.md` (PR #8 — feat/linux-flatpak-build)

**Review:** CodeRabbit review of Linux Flatpak build — 2 actionable findings.
**Result:** Both fixed.

### Findings

1. **`flatpak build-bundle` missing `--runtime-repo`**
   - Bundle was produced without embedding runtime repository metadata; users on clean machines (no Flathub remote) would fail to install if required runtimes were absent
   - Fix: add `--runtime-repo=https://dl.flathub.org/repo/flathub.flatpakrepo` to the `flatpak build-bundle` invocation in CI

2. **README Linux build steps missing Flatpak prerequisites**
   - The documented manual build commands assumed Flathub remote and runtime/SDK were already installed; would fail on a clean system
   - Fix: prepend `flatpak remote-add` and `flatpak install` steps to the README Linux build section; also added `--runtime-repo` to the documented `flatpak build-bundle` command

---

## 2026-04-12 — `modules/config.py` + `modules/app.py` (PR #10 — feat/duplex-reload-instructions)

**Review:** CodeRabbit review of printer-style selector and duplex reload instructions — 2 findings.
**Result:** Both fixed.

### Findings

1. **Legacy `reverse_back` config not migrated to `printer_style`**
   - `load_config()` adopted the new `printer_style` key but never mapped older configs that only contain `reverse_back: false`; those users would silently revert to the face-down default on next launch, losing their preference
   - Fix: in `load_config()`, detect presence of `reverse_back` with absence of `printer_style` and map it: `True → "face_down"`, `False → "face_up"`

2. **UI label strings used as config persistence keys**
   - `printer_style` was saved and compared using the rendered display label (e.g. `"Face-down (most printers)"`); any copy-edit or future localisation would silently break saved preferences and branching logic
   - Fix: introduce `_PRINTER_STYLE_KEYS` (`face_down`/`face_up` → label) and `_PRINTER_STYLE_LABELS` (reverse map) at module level; persist and compare stable keys only; use maps when populating and reading the combo

---

## 2026-04-12 — `build_scripts/HoneyBatchr.iss` + `README.md` (PR #11 — feat/windows-installer)

**Review:** CodeRabbit review of Inno Setup installer — 2 findings.
**Result:** Both fixed.

### Findings

1. **`OutputDir` relative path resolved to wrong directory**
   - `OutputDir=installer_out` in `build_scripts/HoneyBatchr.iss` resolves relative to the script's own directory, emitting to `build_scripts/installer_out`; the CI workflow expected the artifact at repo-root `installer_out/`
   - Fix: change to `OutputDir=..\installer_out` to resolve correctly from the repo root

2. **Windows build snippet used wrong shell dialect label**
   - README code fence was labeled `bash` but contained Windows `set RELEASE_VERSION=dev` syntax
   - Fix: change fence label to `powershell` and use `$env:RELEASE_VERSION = "dev"`

---

## 2026-05-08 — `main.py` + `modules/updater.py` + `modules/app.py` (PR #13 — feat/update-checker)

**Review:** CodeRabbit review of in-app update checker — 3 actionable + 2 nitpicks.
**Result:** All 5 fixed.

### Findings

1. **`flatpak_dns_fix()` called at module top level**
   - Side effect on any import of `main` (tests, packaging tools); should only run when app starts
   - Fix: move call inside `main()`, before `QApplication(sys.argv)`

2. **`random.randint` used for DNS query ID — not cryptographically unpredictable**
   - Predictable QID allows DNS spoofing; response QID was also never validated
   - Fix: replace with `secrets.randbelow(65536)`; validate response length ≥ 12 bytes and QID match before parsing

3. **`_getaddrinfo` wrapper parameter named `type` shadows builtin**
   - Fragile inside function body; Ruff A002 flag
   - Fix: rename parameter to `socktype` in signature and all call sites

4. **Manual "Check for Updates" blocked by `updates_skipped_version` filter**
   - `check_for_updates()` connected to `_on_update_available` which silently returns for skipped versions; user-initiated checks should always show the dialog
   - Fix: add `_on_update_available_manual` that skips the version filter; wire manual checker to it

5. **PE/MZ header check insufficient for installer authenticity (pending)**
   - Format check confirms the file is a Windows executable but does not verify cryptographic integrity or origin
   - Status: logged for future investigation — requires determining whether releases use Authenticode signing or a published SHA-256 digest

---

## 2026-05-08 — `modules/updater.py` + `modules/app.py` (PR #13 second review — feat/update-checker)

**Review:** CodeRabbit follow-up review of fix commit — 2 new findings.
**Result:** Both fixed.

### Findings

1. **`UpdateChecker.run()` silently swallows all network/JSON errors**
   - Broad `except ... pass` meant the UI never learned of failures; manual "Check for Updates" would just do nothing on network error
   - Fix: add `check_failed = pyqtSignal(str)` to `UpdateChecker`; emit it with `str(exc)` instead of `pass`; connect it in `check_for_updates()` (manual path only) to show a `QMessageBox.warning`

2. **`UpdateDownloader.finished` shadows `QThread`'s built-in `finished()` signal**
   - `QThread` emits a parameterless `finished()` when `run()` returns; declaring `finished = pyqtSignal(str)` at the Python level shadows it, causing `thread.finished.connect(worker.deleteLater)` to connect to the wrong overload
   - Fix: rename signal to `download_finished`; update `self.finished.emit(...)`, `downloader.finished.connect(_on_done)`, and `downloader.finished.connect(downloader.deleteLater)` to use the new name

---

## 2026-05-11 — `modules/printing.py` (PR #32 follow-up — feat/progress-bar-cancel)

**Review:** CodeRabbit follow-up — 1 nitpick.
**Result:** Fixed.

### Findings

1. **Exception class defined before imports**
   - `PrintCanceledError` was defined between the stdlib imports and the third-party imports, violating PEP 8
   - Fix: move class definition to after all imports

---

## 2026-05-11 — `modules/app.py` + `modules/printing.py` (PR #32 — feat/progress-bar-cancel)

**Review:** CodeRabbit review of async print worker with progress/cancel — 2 actionable + 1 nitpick.
**Result:** All 3 fixed.

### Findings

1. **`compose_nup_pdf` could not be canceled mid-compose**
   - `PrintWorker.run()` called `compose_nup_pdf` with no cancel callback; a long composition blocked the thread with no way to abort
   - Fix: add `cancel_check=None` param to `compose_nup_pdf`; check it at the start of each sheet iteration and raise `PrintCanceledError` (closing `src` and `out` first); pass `cancel_check=lambda: self._canceled` from `PrintWorker.run()`

2. **`page_progress` callback called before page render completes**
   - `page_progress(i + 1, total_pages)` ran before `page = doc[i]` and the render pipeline; progress was over-reported if render failed
   - Fix: move `page_progress` call to after `painter.drawImage(...)` so it fires only on successful render

3. **Progress dialog lagged one step behind (nitpick)**
   - `_on_step` used `dlg.setValue(current - 1)`, making the bar appear stalled until the next step
   - Fix: change to `dlg.setValue(current)` for accurate real-time progress

---

## 2026-05-10 — `modules/app.py` (PR #31 — feat/recent-files-list)

**Review:** CodeRabbit review of recent files list feature — 1 critical + 1 nitpick.
**Result:** Both fixed.

### Findings

1. **`save_config` erased the recent files list on every save**
   - `save_config` built a fresh dict from UI widgets and called `write_config(data)`, which completely replaced the config file; `recent_files` was not included so every Ctrl+S wiped the list
   - Fix: add `"recent_files": self.config.get("recent_files", [])` to the data dict in `save_config` before calling `write_config`

2. **Clicking a missing recent file gave no feedback**
   - `add_files_to_list` silently skips paths where `os.path.isfile` returns False; user sees nothing happen
   - Fix: add `_add_recent_file(path)` helper — checks existence, shows `QMessageBox.warning`, removes the dead path from the recent list, then calls `add_files_to_list`; wire `_populate_recent_menu` to use this helper
