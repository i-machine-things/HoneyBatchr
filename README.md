# Honey Batchr

A batch printing application built with PyQt6. Matches the Adobe Acrobat Batch Print layout with N-up composition, per-file page configuration, live PDF preview, and theme switching. Runs on Windows and Linux.

---

## Features

- **Batch file queue** — table view with Number, Name, Modified, Range, Copies, Size, Location, State columns
- **Drag & drop** — drop files directly onto the file table
- **N-up printing** — compose 1, 2, 4, 6, 8, 9, or 16 pages per sheet with configurable page order and margins
- **Duplex printing** — hardware duplex (flip on long or short edge via QPrinter) or **manual duplex** (two-pass: prints front sides, prompts to reload paper, prints back sides with optional stack reversal for face-down trays)
- **Orientation** — Portrait, Landscape, or Auto (detects majority orientation of pages to be printed); Auto-Rotate fits each page to its N-up cell; Auto-Center preserves aspect ratio with whitespace padding
- **Paper size** — Letter, Legal, Tabloid, A3, A4, A5; configurable per-side page margins (inch) via Page Setting dialog with live preview
- **Grayscale / color** — per-job color mode
- **Per-file page configuration** — page range (`1,5-9,12`), odd/even subset, per-file copies, reverse pages, live PDF preview with N-up sheet navigation
- **PDF page count** — automatically reads page count from PDFs and shows range in the table
- **Print engine** — PyMuPDF composes the N-up layout; QPrinter/QPainter sends it to the driver (duplex and color applied correctly without `SetPrinter`)
- **Office/other formats** — Word, Excel, PowerPoint fall back to `ShellExecute printto`
- **Theme switcher** — Settings > Theme: Fusion Light, Fusion Dark, Windows, WindowsVista (saved to config, applied before window opens)
- **Settings persistence** — all print options saved to `%USERPROFILE%\.honeybatchr\config.json`
- **Move up / Move down** — reorder files in the queue
- **Context menu** — right-click a file row to remove, open, or open containing folder

---

## Requirements

- Windows 10/11 or Linux (CUPS required for non-PDF file printing on Linux)
- Python 3.11+

### Python dependencies

```
PyQt6>=6.4.0
PyInstaller>=5.0.0
Pillow>=9.0.0
pywin32>=306          # Windows only — installed automatically on Windows
pypdf>=4.0.0
PyMuPDF>=1.23.0
```

Install all:

```bash
pip install -r requirements.txt
```

---

## Running

```bash
python main.py
```

Pass files as arguments to pre-load them (used by the Windows context menu):

```bash
python main.py "C:\docs\file1.pdf" "C:\docs\file2.pdf"
```

---

## Building

**Windows:**
```powershell
python -m PyInstaller build_scripts/HoneyBatchr.spec --clean --noconfirm
$env:RELEASE_VERSION = "dev"
iscc build_scripts\HoneyBatchr.iss
```
Produces `installer_out\HoneyBatchr-dev-windows-setup.exe` via PyInstaller + Inno Setup.
Inno Setup is available at https://jrsoftware.org/isinfo.php.

**Linux:**
```bash
python -m PyInstaller build_scripts/HoneyBatchr.spec --clean --noconfirm
flatpak remote-add --user --if-not-exists flathub https://dl.flathub.org/repo/flathub.flatpakrepo
flatpak install --user -y flathub org.freedesktop.Platform//24.08 org.freedesktop.Sdk//24.08
flatpak-builder --user --install --force-clean flatpak-build build_scripts/com.honeybatchr.HoneyBatchr.yml
flatpak build-bundle ~/.local/share/flatpak/repo HoneyBatchr-linux.flatpak com.honeybatchr.HoneyBatchr \
  --runtime-repo=https://dl.flathub.org/repo/flathub.flatpakrepo
```
Produces a `.flatpak` bundle (Flatpak wrapping the PyInstaller one-dir build).

CI builds both platforms on every version tag and attaches both artifacts to the GitHub release.

---

## Print Engine

| File type | Windows | Linux |
|-----------|---------|-------|
| PDF, XPS, EPUB, CBZ, SVG | PyMuPDF → N-up compose → QPrinter/QPainter | same |
| Images (JPG, PNG, BMP, TIF, GIF) | PyMuPDF → N-up compose → QPrinter/QPainter | same |
| Word, Excel, PowerPoint, other | `ShellExecute printto` (best-effort, N-up not applied) | `lp` via CUPS (best-effort, N-up not applied) |

N-up layout, page borders, and margins are composed by PyMuPDF before the job reaches the printer. Duplex, color mode, and copy count are set through `QPrinter` — no admin rights or `SetPrinter` call needed.

---

## Page Configuration Options

Click **Page Configuration Options...** (or select a file first) to open the per-file dialog:

- **Preview** — live PDF rendering with N-up sheet layout, zoom info, slider navigation
- **Print Range** — All pages, specific range (`1,5-9,12`), odd/even subset
- **Print Specifications** — per-file copies, duplex, reverse pages

Settings are stored on the file entry and applied during composition.

---

## Project Structure

```
HoneyBatchr/
├── main.py                      # Entire application
├── requirements.txt
├── build.bat
├── build_scripts/
│   └── HoneyBatchr.spec         # PyInstaller spec
├── register_context_menu.bat    # Add right-click menu (run as Admin)
├── unregister_context_menu.bat  # Remove right-click menu (run as Admin)
├── create_icons.py              # Generates resources/badger.*
└── resources/
    ├── badger.ico
    ├── badger.png
    └── badger_*.png
```

---

## Configuration file

`%USERPROFILE%\.honeybatchr\config.json`



| Symptom | Fix |
|---------|-----|
| Dark / purple UI on first run | Settings > Theme > Fusion Light |
| PDF preview blank in Page Config dialog | `pip install PyMuPDF` |
| Page count shows "All" instead of range | `pip install pypdf` |
| Office files print 1-up only | Expected — N-up is not applied to ShellExecute path |
| Printer not listed | Check Windows Devices & Printers; restart app |
| Context menu missing | Run `register_context_menu.bat` as Administrator |

