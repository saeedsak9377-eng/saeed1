# How to Build STGS.exe — Complete Step-by-Step Guide

## What you need (one-time setup)

- **Windows 10/11** (must build on Windows to get a Windows .exe)
- **Python 3.10, 3.11, or 3.12** installed — https://www.python.org/downloads/
  - During install: tick **"Add Python to PATH"**
- A folder containing: `stgs_unified.py` and `stgs.spec`

---

## Step 1 — Install dependencies

Open **Command Prompt** (press `Win+R`, type `cmd`, press Enter) and run:

```cmd
pip install pyinstaller pandas numpy openpyxl matplotlib xlsxwriter
```

To also compress the .exe by ~30%, install UPX:

```cmd
winget install upx.upx
```

Or download UPX manually from https://upx.github.io and place `upx.exe` anywhere
on your PATH (e.g. `C:\Windows\System32\`).

---

## Step 2 — Navigate to your project folder

```cmd
cd "C:\path\to\your\project"
```

Example:

```cmd
cd "E:\مكون اختبارات آالي"
```

Verify both files exist:

```cmd
dir stgs_unified.py stgs.spec
```

---

## Step 3 — Build the .exe

Run this single command:

```cmd
pyinstaller stgs.spec
```

This will:
- Analyse all imports automatically
- Bundle Python + all libraries into one file
- Produce `dist\STGS.exe`

The build takes **3–8 minutes** depending on your machine. You will see
a long list of messages — this is normal.

---

## Step 4 — Find your executable

After the build completes:

```
your-project\
    dist\
        STGS.exe        ← this is the file to distribute
    build\              ← temporary files, can be deleted
    stgs.spec
    stgs_unified.py
```

Double-click `dist\STGS.exe` to test it.

---

## Step 5 — Test before distributing

1. Double-click `STGS.exe` — the ETEC launcher should open
2. Try uploading a real bank Excel file
3. Generate 2 forms and verify the output files are created
4. Check that charts appear in the Excel output

**Test on a clean machine** (no Python installed) if possible:
- Copy `STGS.exe` to a USB drive or send via email
- Run it on another Windows PC — it should work with no installation

---

## Troubleshooting

### "Failed to execute script" or blank window

Run from Command Prompt to see the error:

```cmd
dist\STGS.exe
```

### Missing module error (e.g. `ModuleNotFoundError: No module named 'openpyxl'`)

Add the module name to the `hidden` list in `stgs.spec`, then rebuild:

```cmd
pyinstaller stgs.spec
```

### Antivirus flags the .exe as suspicious

This is a **false positive** — very common with PyInstaller executables.
Solutions:
- Add an exception in Windows Defender: Settings → Virus & threat protection → Exclusions
- Or submit to VirusTotal.com to confirm it is clean

### The .exe is very large (200–400 MB is normal)

This is expected — the entire Python runtime and all libraries are bundled.
To reduce size:
1. Ensure UPX is installed (adds `upx=True` compression automatically)
2. The `excludes` list in `stgs.spec` already removes unused heavy packages

### Arabic text or fonts look wrong

Add this to the top of `stgs_unified.py` (already handled by Tkinter on Windows):

```python
import locale
locale.setlocale(locale.LC_ALL, '')
```

### File paths break inside the .exe

If you add extra resource files (icons, config), use this helper instead of
hardcoded paths:

```python
import sys, os

def resource_path(relative_path):
    """Works both when run normally and when bundled by PyInstaller."""
    if hasattr(sys, '_MEIPASS'):
        base = sys._MEIPASS          # PyInstaller temp folder
    else:
        base = os.path.dirname(__file__)
    return os.path.join(base, relative_path)
```

Usage: `resource_path("stgs_icon.ico")`

---

## Optional: Add a custom icon

1. Create or download a 256×256 `.ico` file (Windows icon format)
2. Save it as `stgs_icon.ico` in the same folder as `stgs.spec`
3. In `stgs.spec`, uncomment this line:
   ```python
   # icon='stgs_icon.ico',
   ```
   becomes:
   ```python
   icon='stgs_icon.ico',
   ```
4. Rebuild: `pyinstaller stgs.spec`

---

## Quick rebuild after code changes

Whenever you edit `stgs_unified.py`:

```cmd
pyinstaller stgs.spec
```

The spec file remembers all settings, so you never need to reconfigure.

---

## Distributing to users

Give users only this one file:

```
STGS.exe  (from the dist\ folder)
```

**No Python, no pip, no VS Code, nothing else needed.**

Users just:
1. Copy `STGS.exe` to any folder (e.g. Desktop)
2. Double-click to run
3. Upload their Excel bank and generate tests immediately

---

## Summary of commands

```cmd
# One-time setup
pip install pyinstaller pandas numpy openpyxl matplotlib xlsxwriter

# Build (run from project folder every time you want a new .exe)
cd "path\to\project"
pyinstaller stgs.spec

# Test
dist\STGS.exe
```
