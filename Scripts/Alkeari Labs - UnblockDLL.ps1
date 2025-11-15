<# 
UnblockDLL.ps1
Purpose:
- Removes the Zone.Identifier stream (unblocks) from downloaded files in selected folders.
- Uses the same WinForms OpenFileDialog style as the launcher; no alternate pickers.
- Allows selecting folders or files; file selections are mapped to their parent folders.
- Supports multi-selection across different locations.
- Offers a prompt to either unblock ALL files in the selected folders or only a configured set of extensions.
- Picker windows are TopMost so they remain visible.
- No dry-run; results are logged.

Sections:
1) Environment
2) Picker and target-folder collection (OpenFileDialog; folder or file; multi-select; TopMost)
3) Unblock mode selection (extensions-only vs ALL files)
4) Scan and unblock logic
5) Summary and logging
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Try to import shared theming module if present
$themingModule = Join-Path $PSScriptRoot 'Alkeari.AlkTheming.psm1'
if (Test-Path -LiteralPath $themingModule) {
    Import-Module $themingModule -ErrorAction SilentlyContinue
}

# Fallback console theme helper if module not present
if (-not (Get-Command -Name New-AlkeariConsoleTheme -ErrorAction SilentlyContinue)) {
    function New-AlkeariConsoleTheme {
        [CmdletBinding()]
        param()
        [pscustomobject]@{
            Info   = 'White'
            Warn   = 'Yellow'
            Error  = 'Red'
            Accent = 'Cyan'
        }
    }
}

function Start-AlkeariUnblockDll {
    [CmdletBinding()]
    param()

    $theme = New-AlkeariConsoleTheme

    # =============================
    # 1) Environment (no elevation prompt)
    # =============================

    try {
        Add-Type -AssemblyName System.Windows.Forms | Out-Null
        Add-Type -AssemblyName System.Drawing     | Out-Null
    } catch {
        Write-Host "[ERROR] Could not load required Windows Forms assemblies." -ForegroundColor $($theme.Error)
        Read-Host "Press Enter to close this window..." | Out-Null
        return
    }

    # Cache folder for state and logs
    $CacheRoot = Join-Path $env:TEMP 'Script Cache'
    if (-not (Test-Path -LiteralPath $CacheRoot)) { New-Item -ItemType Directory -Path $CacheRoot | Out-Null }

    # Remember last folder across runs
    $StateFile = Join-Path $CacheRoot 'UnblockDLL.state.json'
    $InitialDir = $env:USERPROFILE
    if (Test-Path -LiteralPath $StateFile) {
        try { $InitialDir = (Get-Content -LiteralPath $StateFile -Raw | ConvertFrom-Json).InitialDir } catch {}
    }
    if (-not $InitialDir -or -not (Test-Path -LiteralPath $InitialDir)) {
        $InitialDir = $env:USERPROFILE
    }

    # Default target extensions (used when not unblocking ALL files)
    $TargetExtensions = @('.dll', '.exe', '.asi', '.ocx', '.ax')  # case-insensitive

    # Log files
    $Timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $LogFile   = Join-Path $CacheRoot ("UnblockDLL-{0}.log" -f $Timestamp)
    $CsvFile   = Join-Path $CacheRoot ("UnblockDLL-{0}.csv" -f $Timestamp)

    # =============================
    # 2) Picker and target-folder collection
    # =============================

    function Resolve-SelectionToFolder {
        param(
            [Parameter(Mandatory)][string]$Selection,
            [Parameter(Mandatory)][string]$Placeholder
        )
        try {
            if ([string]::IsNullOrWhiteSpace($Selection)) { return $null }
            if (Test-Path -LiteralPath $Selection -PathType Container) { return $Selection }
            if (Test-Path -LiteralPath $Selection -PathType Leaf) {
                $dir = [System.IO.Path]::GetDirectoryName($Selection)
                if ($dir -and -not [string]::IsNullOrWhiteSpace($dir)) { return $dir }
            }
            if ($Selection.EndsWith([System.IO.Path]::Combine('', $Placeholder).TrimStart('\\'))) {
                $dir = [System.IO.Path]::GetDirectoryName($Selection)
                if ($dir -and -not [string]::IsNullOrWhiteSpace($dir)) { return $dir }
            }
            $parent = [System.IO.Path]::GetDirectoryName($Selection)
            if ($parent -and -not [string]::IsNullOrWhiteSpace($parent)) { return $parent }
        } catch {}
        return $null
    }

    $folders = New-Object System.Collections.Generic.List[string]
    $PlaceholderName = 'Select Folder'

    # Create a TopMost owner form once and reuse for all dialogs
    $owner = New-Object System.Windows.Forms.Form
    $owner.TopMost = $true
    $owner.ShowInTaskbar = $false
    $owner.FormBorderStyle = 'FixedToolWindow'
    $owner.StartPosition = 'CenterScreen'
    $owner.Size = New-Object System.Drawing.Size(1,1)
    $owner.Opacity = 0.01
    $owner.Show()
    $owner.BringToFront()
    $owner.Activate()

    while ($true) {
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Title = 'Select folders or files. Parent folders will be processed.'
        $dlg.InitialDirectory = $InitialDir
        $dlg.Filter = 'All items (*.*)|*.*'
        $dlg.Multiselect = $true
        $dlg.CheckFileExists = $false
        $dlg.ValidateNames = $false
        $dlg.FileName = $PlaceholderName

        $result = $dlg.ShowDialog($owner)
        if ($result -ne [System.Windows.Forms.DialogResult]::OK) { break }

        if ($dlg.FileNames -and $dlg.FileNames.Count -gt 0) {
            foreach ($sel in $dlg.FileNames) {
                $target = Resolve-SelectionToFolder -Selection $sel -Placeholder $PlaceholderName
                if (-not [string]::IsNullOrWhiteSpace($target) -and (Test-Path -LiteralPath $target -PathType Container)) {
                    $folders.Add($target)
                    $InitialDir = $target
                }
            }
        }

        $more = [System.Windows.Forms.MessageBox]::Show(
            "Add more locations?",
            "Add More",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($more -ne [System.Windows.Forms.DialogResult]::Yes) { break }
    }

    # Dispose owner
    $owner.Close()
    $owner.Dispose()

    # De-duplicate and validate
    $TargetFolders = $folders |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Sort-Object -Unique |
        Where-Object { Test-Path -LiteralPath $_ -PathType Container }

    if (-not $TargetFolders -or $TargetFolders.Count -eq 0) {
        Write-Host "[INFO] No folders selected. Nothing to do." -ForegroundColor $($theme.Info)
        Read-Host "Press Enter to close this window..." | Out-Null
        return
    }

    # Persist last-used directory
    try {
        @{ InitialDir = $InitialDir } |
            ConvertTo-Json |
            Set-Content -LiteralPath $StateFile -Encoding UTF8
    } catch {}

    # =============================
    # 3) Unblock mode selection
    # =============================

    $modeChoice = [System.Windows.Forms.MessageBox]::Show(
        "Do you want to unblock ALL files in the selected folders?" +
        "`n`nYes = Unblock ALL files." +
        "`nNo  = Unblock only these extensions:  {0}" -f ($TargetExtensions -join ', '),
        "Unblock Mode",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    $UnblockAll = ($modeChoice -eq [System.Windows.Forms.DialogResult]::Yes)

    # =============================
    # 4) Scan and unblock logic
    # =============================

    $AllFiles      = New-Object System.Collections.Generic.List[object]
    $Unblocked     = New-Object System.Collections.Generic.List[object]
    $AlreadyClean  = New-Object System.Collections.Generic.List[object]
    $Failed        = New-Object System.Collections.Generic.List[object]

    function Has-ZoneIdentifier {
        param(
            [string]$Path
        )
        if ([string]::IsNullOrWhiteSpace($Path)) {
            return $false
        }
        try {
            $s = Get-Item -LiteralPath $Path -Stream Zone.Identifier -ErrorAction SilentlyContinue
            return $null -ne $s
        } catch { return $false }
    }

    $extHash = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($e in $TargetExtensions) { [void]$extHash.Add($e) }

    # Enumerate
    $folderIndex = 0
    foreach ($root in $TargetFolders) {
        $folderIndex++
        Write-Progress -Activity "Scanning" -Status "$root" -PercentComplete ([math]::Floor(100 * $folderIndex / $TargetFolders.Count))
        try {
            $items = Get-ChildItem -LiteralPath $root -File -Recurse -ErrorAction SilentlyContinue
            foreach ($it in $items) {
                if ($UnblockAll -or $extHash.Contains($it.Extension)) {
                    if (-not [string]::IsNullOrWhiteSpace($it.FullName)) {
                        $AllFiles.Add($it.FullName)
                    }
                }
            }
        } catch {
            $Failed.Add([pscustomobject]@{ Path=$root; Action='Enumerate'; Result='Failed'; Error=$_.Exception.Message })
        }
    }
    Write-Progress -Activity "Scanning" -Completed

    # Deduplicate and validate file list before processing
    $AllFiles = $AllFiles |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and (Test-Path -LiteralPath $_ -PathType Leaf) } |
        Sort-Object -Unique

    if (-not $AllFiles -or $AllFiles.Count -eq 0) {
        Write-Host "[INFO] No matching files found in selected folders." -ForegroundColor $($theme.Warn)
        "`nNo matching files found." | Out-File -LiteralPath $LogFile -Encoding UTF8 -Append
        Read-Host "Press Enter to close this window..." | Out-Null
        return
    }

    # Unblock
    $idx = 0
    foreach ($p in $AllFiles) {
        $idx++
        Write-Progress -Activity "Unblocking" -Status $p -PercentComplete ([math]::Floor(100 * $idx / $AllFiles.Count))
        try {
            if (Has-ZoneIdentifier -Path $p) {
                try {
                    Unblock-File -LiteralPath $p -ErrorAction Stop
                    $Unblocked.Add([pscustomobject]@{ Path=$p; Action='Unblock-File'; Result='Unblocked' })
                } catch {
                    # Fallback: try removing the Zone.Identifier stream directly
                    try {
                        Remove-Item -LiteralPath $p -Stream Zone.Identifier -ErrorAction Stop
                        $Unblocked.Add([pscustomobject]@{ Path=$p; Action='Remove-Item-Stream'; Result='Unblocked' })
                    } catch {
                        $Failed.Add([pscustomobject]@{ Path=$p; Action='Unblock-File/Remove-Stream'; Result='Failed'; Error=$_.Exception.Message })
                    }
                }
            } else {
                $AlreadyClean.Add([pscustomobject]@{ Path=$p; Action='Check'; Result='No Zone.Identifier' })
            }
        } catch {
            $Failed.Add([pscustomobject]@{ Path=$p; Action='Unblock-File'; Result='Failed'; Error=$_.Exception.Message })
        }
    }
    Write-Progress -Activity "Unblocking" -Completed

    # =============================
    # 5) Summary and logging
    # =============================

    $summary = @()
    $summary += "Folders: $($TargetFolders.Count)"
    $summary += "Files matched: $($AllFiles.Count)"
    $summary += "Unblocked: $($Unblocked.Count)"
    $summary += "Already clean: $($AlreadyClean.Count)"
    $summary += "Failed: $($Failed.Count)"
    $summaryText = ($summary -join [Environment]::NewLine)

    $summaryText | Out-File -LiteralPath $LogFile -Encoding UTF8
    "=== Unblocked ===" | Out-File -LiteralPath $LogFile -Encoding UTF8 -Append
    $Unblocked | ForEach-Object { $_.Path } | Out-File -LiteralPath $LogFile -Encoding UTF8 -Append
    "=== Already clean ===" | Out-File -LiteralPath $LogFile -Encoding UTF8 -Append
    $AlreadyClean | ForEach-Object { $_.Path } | Out-File -LiteralPath $LogFile -Encoding UTF8 -Append
    "=== Failed ===" | Out-File -LiteralPath $LogFile -Encoding UTF8 -Append
    $Failed | ForEach-Object { "$($_.Path) :: $($_.Error)" } | Out-File -LiteralPath $LogFile -Encoding UTF8 -Append

    $rows = @()
    $rows += $Unblocked    | ForEach-Object { [pscustomobject]@{ Path=$_.Path; Status='Unblocked';     Error=''          } }
    $rows += $AlreadyClean | ForEach-Object { [pscustomobject]@{ Path=$_.Path; Status='AlreadyClean'; Error=''          } }
    $rows += $Failed       | ForEach-Object { [pscustomobject]@{ Path=$_.Path; Status='Failed';       Error=$_.Error } }
    if ($rows.Count -gt 0) {
        try { $rows | Export-Csv -LiteralPath $CsvFile -NoTypeInformation -Encoding UTF8 } catch {}
    }

    Write-Host "" 
    Write-Host "Folders:        $($TargetFolders.Count)" -ForegroundColor $($theme.Info)
    Write-Host "Files matched:  $($AllFiles.Count)"      -ForegroundColor $($theme.Info)
    Write-Host "Unblocked:      $($Unblocked.Count)"      -ForegroundColor $($theme.Info)
    Write-Host "Already clean:  $($AlreadyClean.Count)"  -ForegroundColor $($theme.Info)
    Write-Host "Failed:         $($Failed.Count)"         -ForegroundColor $($theme.Warn)
    Write-Host "" 
    Write-Host "Log: $LogFile" -ForegroundColor $($theme.Accent)
    if (Test-Path -LiteralPath $CsvFile) { Write-Host "CSV: $CsvFile" -ForegroundColor $($theme.Accent) }

    Read-Host "Press Enter to close this window..." | Out-Null
}

# If the script is executed directly (not dot-sourced), run the entry point
if ($MyInvocation.InvocationName -ne '.') {
    Start-AlkeariUnblockDll
}
