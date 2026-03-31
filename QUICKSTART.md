# Honey Batchr - Quick Start Guide

## What's Included

This package contains everything you need to build and run a batch printing application with PyQt5.

### Core Files

| File | Purpose |
|------|---------|
| `batch_print_app.py` | Main PyQt5 application with all features |
| `create_icons.py` | Generates minimalist badger face icons in multiple sizes |
| `batch_print.spec` | PyInstaller configuration for building the executable |

### Build & Setup Scripts

| File | Purpose |
|------|---------|
| `build.bat` | Main build script (generates icons + builds executable) |
| `setup.bat` | Quick dependency installation |
| `register_context_menu.bat` | Adds "Batch Print with Honey Batchr" to right-click menu (admin required) |
| `unregister_context_menu.bat` | Removes the context menu entry (admin required) |

### Configuration

| File | Purpose |
|------|---------|
| `requirements.txt` | Python package dependencies |
| `README.md` | Comprehensive documentation |
| `QUICKSTART.md` | This file |

---

## Quick Start (3 Steps)

### Step 1: Install Dependencies
```bash
setup.bat
```

### Step 2: Build the Application
```bash
build.bat
```

This will:
- Generate the badger icons
- Build the executable with PyInstaller
- Create `dist\HoneyBatchr\HoneyBatchr.exe`

### Step 3: Run It!
```bash
dist\HoneyBatchr\HoneyBatchr.exe
```

---

## Optional: Add to Context Menu

Run as Administrator:
```bash
register_context_menu.bat
```

Then right-click any file and select **"Batch Print with Honey Batchr"** to add it to the print queue.

---

## Application Features

✅ Drag & drop file support  
✅ Editable file list  
✅ Multiple printer selection  
✅ Copy count, color mode, double-sided options  
✅ Settings persistence  
✅ Minimalist badger icon  
✅ Windows context menu integration  

---

## Troubleshooting

**Q: Build fails**  
A: Run `setup.bat` first to ensure all dependencies are installed

**Q: Context menu not working**  
A: Run `register_context_menu.bat` as Administrator

**Q: Need to customize the icon**  
A: Edit `create_badger_icon()` in `create_icons.py`, then run:
```bash
python create_icons.py
build.bat
```

---

## File Flow

```
setup.bat (install dependencies)
    ↓
build.bat (build executable)
    ├─→ create_icons.py (generate icons)
    └─→ PyInstaller builds from batch_print.spec
        ├─→ batch_print_app.py (main app)
        └─→ resources/ (generated icons)
    ↓
dist/HoneyBatchr/HoneyBatchr.exe (ready to use!)
    ↓
register_context_menu.bat (optional, adds context menu)
```

---

## What the App Does

1. **Add Files**: Click "Add Files", drag-drop, or right-click files (if context menu registered)
2. **Configure**: Choose printer, copies, color mode, etc.
3. **Print**: Click "Print All" to send all files to your printer
4. **Save Settings**: Click "Save Settings" to remember your preferences

---

## Project Structure After Building

```
HoneyBatchr/
├── batch_print_app.py
├── create_icons.py
├── batch_print.spec
├── build.bat
├── register_context_menu.bat
├── unregister_context_menu.bat
├── setup.bat
├── requirements.txt
├── README.md
├── QUICKSTART.md
├── resources/
│   ├── badger.png
│   ├── badger.ico
│   ├── badger_16.png
│   ├── badger_32.png
│   ├── badger_64.png
│   ├── badger_128.png
│   └── badger_256.png
└── dist/
    ├── HoneyBatchr/
    │   ├── HoneyBatchr.exe
    │   ├── launch.bat
    │   └── resources/
    ├── build/
    └── ...
```

---

**Ready to go!** Start with `setup.bat` →  `build.bat` →  run the .exe!
