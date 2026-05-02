# Restore-WindowsFonts

<!-- BADGES:START -->
[![License](https://img.shields.io/github/license/5a9awneh/Restore-WindowsFonts)](LICENSE) [![PowerShell](https://img.shields.io/badge/PowerShell-5391FE?logo=powershell&logoColor=white)](https://learn.microsoft.com/en-us/powershell/) [![Windows](https://img.shields.io/badge/Windows-0078D6?logo=windows&logoColor=white)](https://www.microsoft.com/windows) [![Last Commit](https://img.shields.io/github/last-commit/5a9awneh/Restore-WindowsFonts)](https://github.com/5a9awneh/Restore-WindowsFonts/commits/main) [<img src="https://madebyhuman.iamjarl.com/badges/loop-white.svg" alt="Human in the Loop" height="20">](https://madebyhuman.iamjarl.com)
<!-- BADGES:END -->

Restores corrupted or missing Windows system fonts by extracting them from a Windows ISO you supply, then reinstalling them, clearing font caches, and repairing font registry entries. Resolves Adobe Acrobat "Cannot find or create the font" errors caused by missing Arial, Times New Roman, and Courier New.

## ⚙️ Requirements

- Windows 10 / 11
- PowerShell 5.1 or later, run as Administrator
- Windows installation media *(ISO or USB/DVD — see **Supplying Fonts** below; not needed if `.\.Fonts\` is already populated)*
- DISM PowerShell module *(built into Windows — no third-party tools required)*

## 📁 Supplying Fonts

The script extracts fonts from Windows installation media at runtime. Three options, in order of convenience:

| Option | How |
|--------|-----|
| **USB / DVD** *(easiest)* | Plug in any Windows installation USB or DVD — detected automatically, no extra steps |
| **ISO in script folder** | Drop a `.iso` file into the same folder as the script — picked up automatically |
| **Explicit ISO path** | Pass the path via `RUN.bat "D:\ISOs\Win11.iso"` or `-IsoPath "D:\ISOs\Win11.iso"` |

> Already have the fonts? Populate `.\Fonts\` with the `.ttf` files and the extraction step is skipped entirely.

---

## 🔧 Parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-IsoPath` | `string` | *(auto-detect)* | Full path to a Windows ISO. If omitted, the script looks for a `.iso` file in the script directory. |

## 🚀 Usage

**Via launcher** — handles UAC elevation automatically:

```
RUN.bat
RUN.bat "D:\ISOs\Win11.iso"
```

**Manual** (PowerShell as Administrator):

```powershell
.\Restore-WindowsFonts.ps1
.\Restore-WindowsFonts.ps1 -IsoPath "D:\ISOs\Win11.iso"
```

## 🖼️ Obtaining a Windows 11 ISO *(if needed)*

> Microsoft's download page shows a Media Creation Tool to Windows users. A User-Agent spoof reveals the direct ISO link. **Official source only — do not use third-party sites.**

1. Install a User-Agent switcher extension (Chrome or Firefox)
2. Set your UA to a non-Windows string (e.g., iOS Safari or Android Chrome)
3. Visit https://www.microsoft.com/en-us/software-download/windows11 — the page will now show a direct ISO download instead of the Windows installer
4. Select your language → **64-bit Download**
5. Place the downloaded `.iso` in the script folder, or pass its full path using `-IsoPath`

## 🔍 How It Works

### Step 0 — Font acquisition *(skipped if `.\Fonts\` is already populated)*

1. Checks for the 12 critical fonts (Arial, Times New Roman, and Courier New families)
2. If any are missing, locates Windows installation media in this order:
   - `-IsoPath` parameter value, or `RUN.bat "path\to\windows.iso"`
   - `.iso` file auto-detected in the script folder
   - Windows installation USB drive or DVD (removable/optical drives scanned automatically — no ISO needed)
3. Mounts `install.wim` / `install.esd` via DISM — ISO is mounted first if needed; USB/DVD WIM is accessed directly without mounting
4. Copies all 30 target `.ttf` files from the WIM's `\Windows\Fonts\` into `.\Fonts\`
5. Dismounts the WIM cleanly; dismounts ISO only if it was mounted

### Steps 1–6 — Font repair

1. Stops `FontCache` and `FontCache3.0.0.0` services
2. Clears font cache directories and `FNTCACHE.DAT`
3. Takes ownership of `%WINDIR%\Fonts\*.ttf`
4. Copies all fonts from `.\Fonts\` into `%WINDIR%\Fonts\` via robocopy *(handles locked files)*
5. Restores font registry entries under `HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts`
6. Restarts font services and prompts for a reboot

## ⚠️ Warnings & Limitations

- **Must run as Administrator** — use `RUN.bat` for automatic UAC elevation
- **Reboot required** — changes do not fully apply until Windows restarts
- The 12 critical fonts (Arial, Times New Roman, Courier New families) are required; the remaining 18 (Verdana, Georgia, Tahoma, Calibri, Segoe UI families) are installed if available but non-blocking — the script will warn but continue if they are absent

> **Font sourcing — legal notice:** This repository does not include font files. Arial, Times New Roman, Courier New, Verdana, Georgia, Tahoma, Calibri, and Segoe UI are proprietary fonts distributed with licensed Windows installations and may not be redistributed under their respective Microsoft/Monotype EULAs. The script extracts fonts at runtime from a Windows ISO you supply, sourced from your own licensed copy of Windows. No font files are downloaded from third-party sources.
