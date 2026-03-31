# Honey Batchr - Batch Printing Application

A minimalist PyQt5 application for batch printing files with drag-and-drop support and context menu integration.

## Features

- **Editable File List**: Add, remove, and organize files to print
- **Drag & Drop**: Simply drag files onto the application window
- **Windows Context Menu**: Right-click any file to add it to the print queue
- **Print Settings**: 
  - Printer selection
  - Number of copies
  - Color mode (Color, Grayscale, Monochrome)
  - Double-sided printing
  - Collate option
- **Settings Persistence**: Your preferences are saved locally
- **Minimalist Design**: Clean and simple UI with badger icon

## Installation & Building

### Prerequisites

- Python 3.8 or higher
- Windows OS (for context menu integration)

### Quick Build

Simply run the build script:

```bash
build.bat
```

This will automatically:
1. Check and install dependencies
2. Generate application icons
3. Build the executable with PyInstaller
4. Create the distribution folder

The executable will be created at: `dist\HoneyBatchr\HoneyBatchr.exe`

### Manual Installation

If you prefer to install dependencies manually:

```bash
pip install -r requirements.txt
```

Then build:

```bash
python create_icons.py
pyinstaller batch_print.spec
```

## Usage

### Running the Application

**From Python:**
```bash
python batch_print_app.py
```

**From Executable:**
- Navigate to `dist\HoneyBatchr\` and double-click `HoneyBatchr.exe`
- Or use the `launch.bat` file

### Adding Files to Print

You can add files in three ways:

1. **Click "Add Files"** button to open a file browser
2. **Drag & Drop** files directly onto the application window
3. **Right-click context menu** (after registering - see below)

### Printing

1. Add files to the list
2. Configure print settings:
   - Select printer
   - Set number of copies
   - Choose color mode
   - Enable/disable double-sided printing
   - Set collate option
3. Click **"Print All"** button

The app uses the system's default print handler for each file type.

### Managing the File List

Right-click any file in the list for options to:
- Remove from list
- Open the file
- Open containing folder

## Context Menu Integration

### Register Context Menu (Requires Admin)

After building the application, register it in the Windows context menu:

```bash
register_context_menu.bat
```

**Requirements:** You must run this as Administrator.

This adds "Batch Print with Honey Batchr" to the right-click context menu for all files and folders.

### Unregister Context Menu

To remove the context menu entry:

```bash
unregister_context_menu.bat
```

This must also be run as Administrator.

## Configuration

Settings are automatically saved to: `%USERPROFILE%\.honeybatchr\config.json`

You can also click "Save Settings" in the application to save your current preferences.

## Project Structure

```
HoneyBatchr/
├── batch_print_app.py           # Main application
├── create_icons.py              # Icon generation script
├── batch_print.spec             # PyInstaller specification
├── build.bat                    # Build script
├── register_context_menu.bat    # Install context menu
├── unregister_context_menu.bat  # Remove context menu
├── requirements.txt             # Python dependencies
├── resources/                   # Generated icons
│   ├── badger.png
│   ├── badger.ico
│   └── badger_*.png
└── dist/                        # Built application (after running build.bat)
    └── HoneyBatchr/
        ├── HoneyBatchr.exe      # Executable
        ├── launch.bat           # Launch script
        └── resources/
```

## Supported File Types

The application can print:
- PDF (.pdf)
- Word documents (.doc, .docx)
- Excel spreadsheets (.xls, .xlsx)
- PowerPoint presentations (.ppt, .pptx)
- Images (.jpg, .jpeg, .png, .bmp)
- Text files (.txt)
- And most other file types Windows can print

## Troubleshooting

### Build fails with missing dependencies
Run: `build.bat` - it will automatically install any missing packages

### Context menu not appearing
Make sure to run `register_context_menu.bat` as Administrator

### Application crashes
Check that PyQt5 is properly installed:
```bash
python -c "from PyQt5.QtWidgets import QApplication; print('PyQt5 OK')"
```

### Printing doesn't work
- Make sure your printer is connected and set as default (or select it from the app)
- Try opening the file manually to confirm it prints

## Development

To modify the application:

1. Edit `batch_print_app.py` for application logic
2. Edit `create_icons.py` to change the badger icon design
3. Edit `batch_print.spec` for PyInstaller configuration
4. Run `build.bat` to rebuild

### Icon Customization

To change the badger icon design, edit the `create_badger_icon()` function in `create_icons.py`:

```python
def create_badger_icon(size=256, output_dir="resources"):
    # Modify drawing code here
```

Then regenerate icons with:
```bash
python create_icons.py
```

## License

Free to modify and distribute.

## Notes

- The application uses Windows system print handlers
- Files are sent to the default printer or selected printer
- Printer settings may vary depending on printer capability
- The context menu integration is Windows-specific

---

**Honey Batchr** - Making batch printing honey-sweet! 🦡
