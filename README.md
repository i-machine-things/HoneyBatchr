# Honey Batchr

A Windows batch printing application built with PyQt6. Matches the Adobe Acrobat Batch Print layout with N-up composition, per-file page configuration, live PDF preview, and theme switching.

---

## Features

- **Batch file queue** — table view with Number, Name, Modified, Range, Copies, Size, Location, State columns
- **Drag & drop** — drop files directly onto the file table
- **N-up printing** — compose 1, 2, 4, 6, 8, 9, or 16 pages per sheet with configurable page order and margins
- **Duplex printing** — flip on long edge or short edge via QPrinter (no admin rights required)
- **Grayscale / color** — per-job color mode
- **Per-file page configuration** — page range (`1,5-9,12`), odd/even subset, per-file copies, duplex override, reverse pages, live PDF preview with N-up sheet navigation
- **PDF page count** — automatically reads page count from PDFs and shows range in the table
- **Print engine** — PyMuPDF composes the N-up layout; QPrinter/QPainter sends it to the driver (duplex and color applied correctly without `SetPrinter`)
- **Office/other formats** — Word, Excel, PowerPoint fall back to `ShellExecute printto`
- **Theme switcher** — Settings > Theme: Fusion Light, Fusion Dark, Windows, WindowsVista (saved to config, applied before window opens)
- **Settings persistence** — all print options saved to `%USERPROFILE%\.honeybatchr\config.json`
- **Move up / Move down** — reorder files in the queue
- **Context menu** — right-click a file row to remove, open, or open containing folder

---

## Requirements

- Windows 10/11
- Python 3.11+

### Python dependencies

```
PyQt6>=6.4.0
PyInstaller>=5.0.0
Pillow>=9.0.0
pywin32>=306
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

```bash
build.bat
```

Produces `dist\HoneyBatchr.exe` via PyInstaller.

---

## Print Engine

| File type | Render path |
|-----------|-------------|
| PDF, XPS, EPUB, CBZ, SVG | PyMuPDF → N-up compose → QPrinter/QPainter |
| Images (JPG, PNG, BMP, TIF, GIF) | PyMuPDF → N-up compose → QPrinter/QPainter |
| Word, Excel, PowerPoint, other | `ShellExecute printto` (best-effort, N-up not applied) |

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
├── batch_print.spec             # PyInstaller spec
├── build.bat
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

```json
{
  "copies": 1,
  "collate": true,
  "grayscale": false,
  "print_as_image": false,
  "bleed_marks": false,
  "duplex": false,
  "auto_rotate": true,
  "auto_center": true,
  "orientation": "Portrait",
  "print_what": "Document and markups",
  "simulate_overprint": false,
  "pages_per_sheet": 1,
  "page_order": "Horizontal",
  "margins": 0.2,
  "margins_enabled": true,
  "print_page_border": true,
  "theme": "Fusion Light"
}
```

---

## Troubleshooting

| Symptom | Fix |
|---------|-----|
| Dark / purple UI on first run | Settings > Theme > Fusion Light |
| PDF preview blank in Page Config dialog | `pip install PyMuPDF` |
| Page count shows "All" instead of range | `pip install pypdf` |
| Office files print 1-up only | Expected — N-up is not applied to ShellExecute path |
| Printer not listed | Check Windows Devices & Printers; restart app |
| Context menu missing | Run `register_context_menu.bat` as Administrator |

---

## TODO

### In progress
- [ ] Printer orientation: respect Auto-Rotate and Auto-Center settings during composition
- [ ] Paper size selection (currently hard-coded to US Letter 8.5 x 11)

### Planned
- [ ] Per-file duplex override (currently only global duplex is applied)
- [ ] Booklet print mode — fold-order page imposition
- [ ] Scale mode — fit/fill/percentage scale control
- [ ] Tile Large Pages mode — split oversized pages across sheets
- [ ] Print What: Document only vs. Document and markups (PDF annotation handling)
- [ ] Simulate Overprinting (Fusion blend mode rendering)
- [ ] Bleed Marks output
- [ ] Page Setting dialog (paper size, source tray)
- [ ] Loop printing in list order (continuous/kiosk mode)
- [ ] Windows right-click context menu auto-registration on first run (without requiring Admin separately)
- [ ] Progress bar / cancel button during long print jobs
- [ ] Print preview for the full job before sending
- [ ] Collate support for the GDI print path (currently honoured only by driver for ShellExecute)
- [ ] Recent files list

### Known issues
- [ ] `SetPrinter` raises "Access is denied" — replaced by QPrinter, no longer called
- [ ] Page Config dialog right panel shifts slightly on first show (layout settles after first render)
