# Coding Best Practices & Reminders

> **Style rule:** Notes must be clear and concise — 300 characters or less each. Group by topic, not by date. Whenever a PR review (CodeRabbit or human) catches a mistake, add or amend a note here right away so it isn't repeated.

## Git, CI & Build

- **Match git commit invocations broadly.** Use `\bgit\b.*\bcommit\b` in commit-detecting hooks/regex, not `git\s+commit` — the narrow form misses `git -c flag=val commit`.
- **Keep CI-required build files out of gitignored dirs.** A PyInstaller `.spec` inside a gitignored `build/` breaks CI; keep files the pipeline needs (e.g. in `build_scripts/`) tracked.
- **`flatpak build-bundle` needs `--runtime-repo`.** Without it, bundles fail to install on clean machines that lack the Flathub remote/runtime already configured.
- **Inno Setup `OutputDir` is relative to the `.iss` file, not the repo root.** Use `..\installer_out` if CI expects the artifact at repo-root `installer_out/`.

## Config & Persistence

- **Never track machine-specific config.** Gitignore files like `settings.local.json` that contain absolute paths/wildcards; ship a `.example` with minimal portable entries instead.
- **Migrate legacy config keys, don't just add new ones.** Renaming a key (e.g. `reverse_back` → `printer_style`) needs a load-time mapping or existing users silently lose their saved preference.
- **Persist stable keys, never UI display labels.** Saving a setting as its rendered label text breaks on any copy-edit or localization; use a key↔label map and persist/compare only the key.
- **Config saves must round-trip fields the current UI doesn't own.** Rebuilding the save dict from visible widgets alone silently drops other keys (e.g. `recent_files`) on every save.

## Printing & PDF Composition

- **Detect layout only from pages that will actually print.** Scanning every page while ignoring active print-range/odd-even/reverse filters causes wrong orientation detection; filter indices first.
- **Prefer library-native scaling over reimplementing it.** Manual aspect-ratio math can duplicate what `show_pdf_page(keep_proportion=True)` already does — pass the flag instead of recomputing scale.
- **Validate numeric config before unit conversion.** Check domain constraints (e.g. margins `>= 0`) before converting to points, and raise a clear error naming the offending value.
- **Wire new dialog settings all the way to their call sites.** A setting stored in config is useless if the composer still reads old widget state — verify every new option actually reaches the code that uses it.
- **Validate combined values against the real limit, not just each part.** Bounding each margin spinbox individually isn't enough; left+right or top+bottom can still exceed the sheet dimension.
- **Job-wide options must apply to every affected file type.** A duplex/print-mode toggle that only touches one file type (e.g. PDFs) should reject or explicitly handle the unsupported types, not silently skip them.

## Qt / UI

- **Create widgets before referencing them during init.** Calling a handler that touches a widget before that widget is constructed raises `AttributeError` if the feature starts enabled; order creation before wiring.
- **User-cancelled multi-step actions must not report success.** If a mid-flow prompt (e.g. reload-for-back-side) is cancelled, track that state and message accordingly instead of showing the normal success text.
- **Don't name a custom signal the same as a Qt base-class signal.** Declaring `finished = pyqtSignal(str)` shadows `QThread`'s built-in parameterless `finished()`, misconnecting callers. Use a distinct name.
- **Give feedback when an action silently no-ops.** A click that fails a hidden precondition (e.g. missing file) should show a warning and clean up the stale entry, not do nothing visible.

## Security

- **Use `secrets`, not `random`, for anything an attacker could predict.** Predictable IDs (e.g. DNS query IDs from `random.randint`) enable spoofing; use `secrets.randbelow` and validate responses (length, ID match).
- **A file-format check is not an authenticity check.** Confirming a PE/MZ header proves a downloaded file is an executable, not that it's genuine — still need signing verification or a published checksum.

## Error Handling & Cleanup

- **Catch specific exceptions, not bare `Exception`.** Broad catches around subprocess calls or JSON parsing hide real failures; use targeted exception types and check return codes explicitly.
- **Clean up temp files on every failure path.** Wrap creation/use in try/finally (or explicit except+unlink) so a raise before the try block, a later step failing, or a save failure doesn't leak temp files.
- **Background workers must surface errors to the UI.** A bare `except: pass` in a `QThread.run()` makes failures invisible; emit an error signal and connect it to a warning dialog.

## Code Quality

- **Don't shadow builtins in parameter names.** A parameter named `type` trips linters (e.g. Ruff A002) and is fragile inside the function body; use a specific name like `socktype`.
- **Annotate return types precisely.** Use `tuple[str, str | None]` instead of a bare `tuple` so IDEs and type-checkers know what's actually returned.

## Docs

- **Document every prerequisite for manual build steps.** Instructions that assume a remote/runtime/SDK is already installed will fail on a clean machine; list the setup steps too.
- **Match code-fence language to the actual shell.** Windows `set`/`$env:` snippets labeled `bash` are confusing; label them `powershell`.

## Feature Behavior

- **User-triggered actions should bypass automatic filters.** A manual "check now" action reusing the background auto-check's handler can silently skip results (e.g. already-dismissed versions); manual actions should always show results.
- **Side-effecting calls belong inside `main()`, not at module import time.** A fix/setup function that runs on every import affects tests and packaging tools; call it inside `main()` before app startup instead.
