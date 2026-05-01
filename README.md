# Starfield Item Codex

**Complete FormID database for every player-obtainable item in Starfield.**

Base Game · Shattered Space · Terran Armada · Trackers Alliance · All Official DLC

> Current as of **CAO Starfield Update 1.16.236** – April 7, 2026

---

<p align="center">
  <img src="https://img.shields.io/badge/platform-Windows%208%2B-blue?style=flat-square" alt="Platform">
  <img src="https://img.shields.io/badge/built%20with-Python%20%2B%20Excel-yellow?style=flat-square" alt="Built With">
  <img src="https://img.shields.io/badge/AstralUI-supported-green?style=flat-square" alt="AstralUI Support">
</p>

---

## What Is This?

A standalone desktop tool for searching, browsing, and exporting Starfield item data. Three tabs:

| Tab | What It Does |
|-----|-------------|
| 🔍 **Search** | Browse and filter the full FormID database |
| 🏗 **Subcategory Builder** | Build custom INI subcategory entries for [AstralUI](https://www.nexusmods.com/starfield/mods/16886) |
| 🎮 **Batch Creator** | Generate console-command `.txt` batch files to spawn items in-game |

---

## Getting Started
Note: The in-code name is AstralCodex, the project and exe name on the nexus is StarfieldItemCodex. Will sort out later, they are the same. 

### Option A: Run the Pre-Built Executable (Recommended)

1. https://www.nexusmods.com/starfield/mods/16886
2. Extract the `.zip` somewhere convenient
3. **Read `HOWTO_VERIFY_SAFE.txt`** - it explains how to verify the `.exe` is clean
4. Make sure `FormID_List.xlsx` is in the **same folder** as `StarfieldItemCodex.exe`
5. Double-click `StarfieldItemCodex.exe`

> **That's it.** No installation, no setup, no dependencies.

### Option B: Run from Source (Python)

If you'd rather run the `.py` directly:

1. Install **Python 3.10+** from [python.org](https://www.python.org/downloads/)
   - Check **"Add Python to PATH"** during install
2. Open a terminal in the project folder and install dependencies:
   ```
   pip install customtkinter openpyxl
   ```
3. Make sure `FormID_List.xlsx` is in the same folder as `AstralCodex.py`
4. Run it:
   ```
   python AstralCodex.py
   ```

---

## Localization

Full UI and item name translations for every language Starfield natively supports:

| | | | |
|---|---|---|---|
| 🇺🇸 English | 🇩🇪 German | 🇪🇸 Spanish | 🇫🇷 French |
| 🇮🇹 Italian | 🇯🇵 Japanese | 🇵🇱 Polish | 🇧🇷 Portuguese (BR) |
| 🇨🇳 Chinese (Simplified) | | | |

Switch languages at any time from the dropdown in the top-right corner. Item names update to match.

---

## Compatibility

- **Windows 8** or newer
- Works on Windows 10/11 out of the box
- For source: Python 3.10+ required

---

## Features

Full feature list and screenshots on Nexus Mods:

**[Starfield Item Codex on Nexus Mods](https://www.nexusmods.com/starfield/mods/16886)**


---

## Credits

**Developed by cybernav**

Built with Python and Excel - because Excel is the best kind of database.

---

## License

This project is licensed under the [MIT License](LICENSE).
