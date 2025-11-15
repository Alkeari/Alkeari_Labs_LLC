# Alkeari Labs LLC - Windows Startup Manager

A Windows-only Avalonia desktop application to view, add, update, and remove startup applications while adhering to the Alkeari Labs LLC brand theme.

## Features
- Enumerates startup items from:
  - HKCU and HKLM Run registry keys
  - User and Common Startup folders (shortcuts resolved to target + arguments)
- Add or update entries in selected location (creates .lnk for Startup folders, writes quoted path to registry keys)
- Remove existing entries
- Brand-compliant color palette and typography (Cinzel for headings, Crimson Text for body)
- Fonts packaged with the app; can be installed per-user without admin rights.

## Requirements
- Windows OS (registry + COM shell shortcut APIs)
- .NET 8 SDK

## Running
```
dotnet build
cd "Windows Startup Manager/bin/Debug/net8.0"
"Windows Startup Manager.exe"
```

## Fonts
- The app bundles Cinzel, Crimson Text, and Roboto Mono as resources and uses them directly without requiring system installation.
- Optional: Click "Install Fonts" in the app header to register fonts for the current Windows user (no admin required). This improves appearance in other apps and ensures system-wide availability for that user.

## Adding a Startup Item
1. Click New.
2. Enter a Name (unique per location), Command (full path to executable or script), optional Arguments.
3. Choose Location.
4. Click Save.

## Removing
1. Select item in list.
2. Click Delete.

## Permissions
- Writing to HKLM Run or Common Startup folder usually needs elevated (Administrator) privileges. If access is denied, run the application as Administrator.

## Notes
- Shortcut creation uses WScript.Shell COM automation; will throw NotSupportedException if unavailable (non-Windows environments are not supported).
- Disabling (without removal) is not yet implemented; future enhancement could toggle entries by moving them to a disabled store or prefixing value.

## Brand Theme
Colors, fonts defined in `App.axaml` per `.AI Assistant/Alkeari Labs LLC - Theme.txt`.

## Next Steps (Potential Enhancements)
- Add enable/disable toggle.
- Support Task Scheduler startup entries.
- Validate paths with file pick dialog.
- Add search/filter bar.
- Persist window size/position.

## License
Internal proprietary project for Alkeari Labs LLC.
