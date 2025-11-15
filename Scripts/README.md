# Alkeari Labs LLC – Shared Scripts

This folder contains supporting PowerShell tools for the Alkeari Labs LLC solution. They are no longer tied to the `Windows Startup Manager` project and can be used independently across Alkeari tooling.

They share a common theming module and follow a consistent pattern:

- Strict mode (`Set-StrictMode -Version Latest`) for safer scripting.
- A single `Start-*` entry function per script.
- Auto-run when executed as a script, but **no auto-run when dot-sourced/imported**.
- No required CLI parameters – all functionality is available via GUI/dialogs plus human-readable console output.

> **Execution policy**: If needed, run PowerShell as Administrator and use:
>
> ```powershell
> Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
> ```

## Shared theming module

### `Alkeari.AlkTheming.psm1`

A small module that centralizes Alkeari-specific console and UI color palettes:

- `New-AlkeariIconBranding()` – console colors for Icon Suite (Gold/Silver/etc.).
- `New-AlkeariFolderTheme()` – WPF color palette for Folder Inventory.
- `New-AlkeariModInstallerTheme()` – colors for Smart Mod Installer.
- `New-AlkeariConsoleTheme()` – generic Info/Warn/Error/Accent console colors.

All tools in this folder try to import this module from the same directory:

```powershell
$themingModule = Join-Path $PSScriptRoot 'Alkeari.AlkTheming.psm1'
if (Test-Path -LiteralPath $themingModule) {
    Import-Module $themingModule -ErrorAction SilentlyContinue
}
```

If the module is missing, each script defines a compatible local fallback so it continues to work.

---

## Tools

### 1. Folder Inventory & File Tree

**Script:** `Alkeari Labs - Folder Inventory & File Tree.ps1`

Generates a YAML report (and optional JSON) plus a text tree view for a selected folder:

- Recursively scans a target directory.
- Applies optional filters:
  - Include hidden.
  - Include system.
- Produces:
  - `FolderReport_<timestamp>.yaml` (always).
  - `FolderReport_<timestamp>.json` (if checked).
  - `FileTree_<timestamp>.txt` (ASCII tree).
- Provides progress and logging inside a WPF window.

**How to run (GUI):**

```powershell
cd "C:\Users\Josep\RiderProjects\Alkeari Labs LLC\Scripts"
.".\Alkeari Labs - Folder Inventory & File Tree.ps1"
```

You’ll see a window titled **“Alkeari Labs LLC – Folder Inventory”**. Choose the target folder, output folder, and options, then click **Generate**.

**Programmatic use (optional):**

```powershell
. ."Alkeari Labs - Folder Inventory & File Tree.ps1"  # dot-source
Start-AlkeariFolderInventory                           # run the tool
```

---

### 2. Icon Suite Generator

**Script:** `Alkeari Labs - Icon Suite Generator.ps1`

Creates multi-size `.ico` files from one or more images.

- Self-bootstraps dependencies into:
  - `%APPDATA%\Alkeari Labs LLC\IconSuite\deps`
- Supported formats (via Magick.NET / Skia / WIC fallback):
  - PNG, JPG/JPEG, BMP, GIF, TIFF, TGA, WEBP, ICO, SVG, HEIC/HEIF, and more.
- For each selected image, renders multiple sizes:
  - `256, 128, 64, 32, 16` pixels.
- Packs all sizes into a single `.ico` using `New-IcoFromPngMap`.
- Writes output into a sibling folder next to each source image:
  - `IconSuite_<timestamp>\<SourceName>.ico`.

**How to run (GUI + console):**

```powershell
cd "C:\Users\Josep\RiderProjects\Alkeari Labs LLC\Scripts"
.".\Alkeari Labs - Icon Suite Generator.ps1"
```

You’ll get a console header with Alkeari branding, then a file picker dialog. Select one or more images; summary output and any warnings are printed to the console.

**Programmatic use (optional):**

```powershell
. ."Alkeari Labs - Icon Suite Generator.ps1"  # dot-source
Start-AlkeariIconSuite                        # run the tool
```

---

### 3. Smart Mod Installer

**Script:** `Alkeari Labs - SmartModInstaller.ps1`

A smart installer for Bannerlord modules that works entirely via dialogs and console output (no CLI flags):

- Remembers last working and target directories in `%LOCALAPPDATA%\BannerlordSmartInstaller\config.xml`.
- Prompts for:
  - Working directory containing mod archives.
  - Target `Modules` directory.
- Supports many archive formats (`.zip`, `.7z`, `.rar`, `.tar`, `.gz`, `.bz2`, `.cab`, etc.) using:
  - Built-in `Expand-Archive` for `.zip`.
  - PeaZip, 7-Zip, or WinRAR when available.
- Installs mods by:
  - Detecting `SubModule.xml` for full installs.
  - Detecting `Modules/<ExistingModuleName>` layouts for patch-style archives.
  - Using per-module mutexes and retries to avoid file-lock issues.
- Moves processed archives to an `Installed` folder in the working directory.

**How to run:**

```powershell
cd "C:\Users\Josep\RiderProjects\Alkeari Labs LLC\Scripts"
.".\Alkeari Labs - SmartModInstaller.ps1"
```

Dialogs will guide you to select the working directory and target Modules directory; progress and results are printed to the console.

**Programmatic use (optional):**

```powershell
. ."Alkeari Labs - SmartModInstaller.ps1"  # dot-source
Start-AlkeariSmartModInstaller             # run the tool
```

---

### 4. UnblockDLL

**Script:** `Alkeari Labs - UnblockDLL.ps1`

Unblocks downloaded files by removing the `Zone.Identifier` stream in selected folders.

- Requests elevation (with a GUI prompt) and restarts itself as Administrator if approved.
- Uses `OpenFileDialog` to select one or more folders or files:
  - File selections are mapped to their parent folders.
  - Multi-select is supported.
- Lets you choose:
  - Unblock **all** files in selected folders, or
  - Only specific extensions (`.dll`, `.exe`, `.asi`, `.ocx`, `.ax`).
- Recursively scans selected folders and runs `Unblock-File` when needed.
- Writes logs to `%TEMP%\Script Cache`:
  - `UnblockDLL-<timestamp>.log`
  - `UnblockDLL-<timestamp>.csv` (optional summary table).

**How to run:**

```powershell
cd "C:\Users\Josep\RiderProjects\Alkeari Labs LLC\Scripts"
.".\Alkeari Labs - UnblockDLL.ps1"
```

The script will prompt for elevation if needed, then show dialogs for selecting locations. A summary is printed to the console, along with log paths.

**Programmatic use (optional):**

```powershell
. ."Alkeari Labs - UnblockDLL.ps1"  # dot-source
Start-AlkeariUnblockDll            # run the tool
```

---

## Troubleshooting

- **Execution policy errors** – if you see messages about scripts being disabled, run PowerShell as Administrator and execute:

  ```powershell
  Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
  ```

- **Missing external tools** – Smart Mod Installer prefers PeaZip/7-Zip/WinRAR for non-zip archives. If no compatible archive tool is installed, some formats may fail to extract.
- **Magick.NET / Skia issues** – Icon Suite Generator bootstraps NuGet packages into `%APPDATA%`. If the first run fails due to network issues, re-run after ensuring internet access, or clear the `IconSuite\deps` folder and try again.
- **File locks** – Smart Mod Installer and UnblockDLL may encounter locked files if the game or launcher is running. Close related applications and try again.
