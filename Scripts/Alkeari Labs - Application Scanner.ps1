<# 
    Alkeari Labs LLC - Application Inventory Scanner
    Description:
        Scans the system for:
          - Installed programs (from registry)
          - Start Menu and Desktop shortcuts
          - Executable files in common application folders

    Notes:
        - Best run in a PowerShell session with appropriate permissions.
        - Full scan can take a while on some systems.

    Adjusted goals:
        - Focused results on user-facing applications, not Windows/system components.
        - De-duplicated entries across registry, shortcuts, and executables.
        - When multiple entries represented the same app, preferred .exe > shortcut > registry-only.
        - Produced a clean list suitable for deciding what to auto-start on Windows boot.
#>

# Clear console and show Alkeari Labs LLC UI header
Clear-Host

# ====================== Alkeari Labs LLC Header ============================
$brandName     = "Alkeari Labs LLC"
$primaryColor  = "Black"
$accentGold    = "Yellow"
$accentBrown   = "DarkYellow"
$neutralSilver = "Gray"
$neutralWhite  = "White"

Write-Host ""
Write-Host "==================================================================" -ForegroundColor $accentGold
Write-Host "    $brandName - Application Inventory Scanner" -ForegroundColor $neutralWhite
Write-Host "==================================================================" -ForegroundColor $accentGold
Write-Host "" 
Write-Host "  This tool focused on user-facing apps and reduced duplicate/system noise" -ForegroundColor $neutralSilver
Write-Host "  so that you could decide what to run on Windows startup." -ForegroundColor $neutralSilver
Write-Host "" 
Write-Host "==================================================================" -ForegroundColor $accentGold
Write-Host "" 

# --------------------------- Helper Functions -----------------------------

function Write-Info {
    param([string]$Message)
    Write-Host "[INFO]  $Message" -ForegroundColor Cyan
}

function Write-Warn {
    param([string]$Message)
    Write-Host "[WARN]  $Message" -ForegroundColor Yellow
}

function Write-ErrorMsg {
    param([string]$Message)
    Write-Host "[ERROR] $Message" -ForegroundColor Red
}

function Write-Ok {
    param([string]$Message)
    Write-Host "[OK]    $Message" -ForegroundColor Green
}

# Normalized label generator used for de-duplication and scoring
function Get-AppIdentityKey {
    param(
        [string]$Name,
        [string]$TargetPath
    )

    $n = ($Name  | ForEach-Object { $_.ToLowerInvariant().Trim() })
    $p = ($TargetPath | ForEach-Object { $_.ToLowerInvariant().Trim() })

    if ([string]::IsNullOrWhiteSpace($n) -and [string]::IsNullOrWhiteSpace($p)) {
        return $null
    }

    # Remove common noise words from name
    $noise = @('setup','installer','update','updater','assistant','helper')
    if ($n) {
        foreach ($w in $noise) {
            $n = $n -replace "\b$w\b",""
        }
        $n = $n -replace '\s+',' '
    }

    # Normalize path by directory + basename when present
    if ($p) {
        try {
            $fileName = [System.IO.Path]::GetFileNameWithoutExtension($p)
            $dirName  = [System.IO.Path]::GetFileName([System.IO.Path]::GetDirectoryName($p))
            if ($fileName) {
                $p = ("$dirName/$fileName").ToLowerInvariant()
            }
        } catch {}
    }

    if ($n -and $p) { return "$n|$p" }
    if ($n)          { return "name|$n" }
    if ($p)          { return "path|$p" }
}

# Rough system noise filter so results focused on end-user apps
function Test-IsSystemApp {
    param(
        [string]$Name,
        [string]$Path
    )

    $n = $Name
    $p = $Path

    if ($n) { $n = $n.ToLowerInvariant() }
    if ($p) { $p = $p.ToLowerInvariant() }

    # Filter obvious Windows / driver / runtime entries
    $systemNameTokens = @(
        'microsoft visual c++',
        'redistributable',
        'runtime',
        'update for ',
        'security update',
        'hotfix',
        'driver',
        'support assistant',
        'bios',
        'firmware',
        'windows sdk',
        'windows software development kit',
        'microsoft .net',
        '.net runtime',
        'vc++',
        'cumulative update'
    )

    foreach ($token in $systemNameTokens) {
        if ($n -and $n.Contains($token)) { return $true }
    }

    # Filter known system locations
    if ($p) {
        $systemPathTokens = @(
            '\\windows\\',
            '\\program files\\windows',
            '\\program files (x86)\\windows',
            '\\windowsapps\\',
            '\\microsoft\\edge',
            '\\internet explorer\\',
            '\\microsoft office\\root\\vfs',
            '\\common files\\microsoft shared'
        )
        foreach ($t in $systemPathTokens) {
            if ($p.Contains($t)) { return $true }
        }
    }

    return $false
}

# Ranking for de-duplication: higher score wins
function Get-AppScore {
    param(
        [string]$Type,
        [string]$Path
    )

    # Base type preference: Exe > Shortcut > Registry entry
    $score = switch ($Type) {
        'Executable'       { 300 }
        'Shortcut'         { 200 }
        'InstalledProgram' { 100 }
        default            { 0 }
    }

    # Slight bonus for paths under Program Files vs user profile
    if ($Path) {
        $pl = $Path.ToLowerInvariant()
        if ($pl -like "$env:ProgramFiles*".ToLowerInvariant() -or
            $pl -like "$env:ProgramFiles(x86)*".ToLowerInvariant()) {
            $score += 20
        } elseif ($pl -like "$env:LocalAppData*".ToLowerInvariant() -or
                  $pl -like "$env:AppData*".ToLowerInvariant()) {
            $score += 10
        }
    }

    return $score
}

# Get installed programs from registry
function Get-InstalledPrograms {
    Write-Info "Scanning registry for installed programs..."

    $paths = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
    )

    $programs = @()

    foreach ($path in $paths) {
        if (Test-Path $path) {
            Get-ChildItem $path -ErrorAction SilentlyContinue | ForEach-Object {
                try {
                    $displayName = $_.GetValue('DisplayName')
                    if ([string]::IsNullOrWhiteSpace($displayName)) { return }

                    $location = $_.GetValue('InstallLocation')
                    $exePath  = $_.GetValue('DisplayIcon')

                    if (Test-IsSystemApp -Name $displayName -Path $location) { return }

                    $programs += [PSCustomObject]@{
                        Type            = 'InstalledProgram'
                        Name            = $displayName
                        Version         = $_.GetValue('DisplayVersion')
                        Publisher       = $_.GetValue('Publisher')
                        InstallDate     = $_.GetValue('InstallDate')
                        InstallLocation = $location
                        Source          = $path
                        Path            = $exePath
                    }
                } catch {
                    # ignore broken entries
                }
            }
        }
    }

    Write-Ok "Found $($programs.Count) installed program entries (pre-filtered)."
    return $programs
}

# Resolve shortcut target using COM
function Resolve-Shortcut {
    param(
        [string]$ShortcutPath
    )

    try {
        $wshShell = New-Object -ComObject WScript.Shell
        $shortcut = $wshShell.CreateShortcut($ShortcutPath)
        return $shortcut.TargetPath
    } catch {
        return $null
    }
}

# Get Start Menu and Desktop shortcuts
function Get-ShortcutEntries {
    Write-Info "Scanning Start Menu and Desktop for shortcuts..."

    $locations = @(
        "$env:ProgramData\Microsoft\Windows\Start Menu\Programs",
        "$env:AppData\Microsoft\Windows\Start Menu\Programs",
        "$env:Public\Desktop",
        "$env:UserProfile\Desktop"
    ) | Where-Object { Test-Path $_ }

    $shortcuts = @()

    foreach ($location in $locations) {
        Write-Info "Scanning shortcuts in: $location"
        Get-ChildItem -Path $location -Filter *.lnk -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
            $target = Resolve-Shortcut -ShortcutPath $_.FullName
            if (-not $target) { return }

            if (Test-IsSystemApp -Name $_.BaseName -Path $target) { return }

            $shortcuts += [PSCustomObject]@{
                Type            = 'Shortcut'
                Name            = $_.BaseName
                Path            = $_.FullName
                TargetPath      = $target
                Source          = $location
                Version         = $null
                Publisher       = $null
                InstallDate     = $null
                InstallLocation = (Split-Path -Path $target -Parent)
            }
        }
    }

    Write-Ok "Found $($shortcuts.Count) shortcut entries (pre-filtered)."
    return $shortcuts
}

# Get executables in common application locations
function Get-ExecutableEntries {
    Write-Info "Scanning for .exe files in common application folders..."
    $paths = @(
        "$env:ProgramFiles",
        "$env:ProgramFiles(x86)",
        "$env:LocalAppData",
        "$env:AppData"
    ) | Where-Object { Test-Path $_ }

    $executables = @()

    foreach ($path in $paths) {
        Write-Info "Scanning executables in: $path"
        try {
            Get-ChildItem -Path $path -Filter *.exe -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
                if (Test-IsSystemApp -Name $_.BaseName -Path $_.FullName) { return }

                $executables += [PSCustomObject]@{
                    Type            = 'Executable'
                    Name            = $_.BaseName
                    Path            = $_.FullName
                    Source          = $path
                    Version         = $null
                    Publisher       = $null
                    InstallDate     = $null
                    InstallLocation = (Split-Path -Path $_.FullName -Parent)
                    TargetPath      = $_.FullName
                }
            }
        } catch {
            Write-Warn "Could not fully scan: $path"
        }
    }

    # Remove duplicates by path
    $executables = $executables | Sort-Object -Property Path -Unique

    Write-Ok "Found $($executables.Count) executable entries (pre-filtered)."
    return $executables
}

# --------------------------- Merge / De-dupe -------------------------------

function Merge-And-DeduplicateApps {
    param(
        [array]$Installed,
        [array]$Shortcuts,
        [array]$Executables
    )

    Write-Info "Merging and de-duplicating application entries..."

    $all = @()
    if ($Installed)   { $all += $Installed }
    if ($Shortcuts)   { $all += $Shortcuts }
    if ($Executables) { $all += $Executables }

    $map = @{}

    foreach ($item in $all) {
        $name = $item.Name
        $path = if ($item.TargetPath) { $item.TargetPath } elseif ($item.Path) { $item.Path } elseif ($item.InstallLocation) { $item.InstallLocation } else { $null }

        if (Test-IsSystemApp -Name $name -Path $path) { continue }

        $key = Get-AppIdentityKey -Name $name -TargetPath $path
        if (-not $key) { continue }

        $score = Get-AppScore -Type $item.Type -Path $path

        if (-not $map.ContainsKey($key)) {
            $map[$key] = @{ Item = $item; Score = $score }
        } else {
            if ($score -gt $map[$key].Score) {
                $map[$key] = @{ Item = $item; Score = $score }
            }
        }
    }

    $result = $map.Values | ForEach-Object { $_.Item }

    # Sort by name for readability
    $sorted = $result | Sort-Object -Property Name

    Write-Ok "Final unique, filtered applications: $($sorted.Count)"
    return $sorted
}

# --------------------------- Main Menu Logic -------------------------------

function Show-MainMenu {
    Write-Host "" 
    Write-Host "Select scan mode:" -ForegroundColor $neutralWhite
    Write-Host "  [1] Quick Scan  (Installed programs + Shortcuts)" -ForegroundColor $neutralSilver
    Write-Host "  [2] Full Scan   (Quick scan + Executables in common folders)" -ForegroundColor $neutralSilver
    Write-Host "  [3] Exit" -ForegroundColor $neutralSilver
    Write-Host "" 
}

function Export-Results {
    param(
        [array]$Results
    )

    if (-not $Results -or $Results.Count -eq 0) {
        Write-Warn "No results to export."
        return
    }

    Write-Host "" 
    Write-Host "Would you like to export the results to an Excel .xlsx file?" -ForegroundColor $neutralWhite
    Write-Host "  [Y] Yes" -ForegroundColor $neutralSilver
    Write-Host "  [N] No" -ForegroundColor $neutralSilver
    $choice = Read-Host "Enter choice (Y/N)"

    if ($choice -notmatch '^[Yy]') {
        Write-Info "Skipping export."
        return
    }

    # Prefer XLSX via ImportExcel if available; fall back to CSV otherwise.
    $hasImportExcel = $false
    try {
        if (Get-Module -ListAvailable -Name ImportExcel -ErrorAction SilentlyContinue) {
            $hasImportExcel = $true
        }
    } catch {}

    if ($hasImportExcel) {
        $defaultPath = Join-Path -Path ([Environment]::GetFolderPath('Desktop')) -ChildPath 'ApplicationInventory-Clean.xlsx'
        $exportPath  = Read-Host "Enter path for Excel file or press Enter for default (`"$defaultPath`")"
        if ([string]::IsNullOrWhiteSpace($exportPath)) {
            $exportPath = $defaultPath
        }

        try {
            # Create a nicely formatted Excel worksheet
            $Results |
                Select-Object Type, Name, Path, TargetPath, Version, Publisher, InstallLocation |
                Export-Excel -Path $exportPath -WorksheetName 'Applications' -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow

            Write-Ok "Results exported to Excel: $exportPath"
        } catch {
            Write-ErrorMsg "Failed to export to Excel. Error: $($_.Exception.Message)"
        }
    } else {
        Write-Warn "Module 'ImportExcel' not found. Falling back to CSV export."

        $defaultPath = Join-Path -Path ([Environment]::GetFolderPath('Desktop')) -ChildPath 'ApplicationInventory-Clean.csv'
        $exportPath  = Read-Host "Enter path for CSV file or press Enter for default (`"$defaultPath`")"
        if ([string]::IsNullOrWhiteSpace($exportPath)) {
            $exportPath = $defaultPath
        }

        try {
            $Results |
                Select-Object Type, Name, Path, TargetPath, Version, Publisher, InstallLocation |
                Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8
            Write-Ok "Results exported to CSV: $exportPath"
        } catch {
            Write-ErrorMsg "Failed to export CSV. Error: $($_.Exception.Message)"
        }
    }
}

# Store results globally for later inspection if desired
$global:ApplicationInventoryResults = @()

$done = $false
while (-not $done) {
    Show-MainMenu
    $selection = Read-Host "Enter your selection (1-3)"

    switch ($selection) {
        '1' {
            Write-Info "Starting Quick Scan (registry + shortcuts)..."
            $installed = Get-InstalledPrograms
            $shortcuts = Get-ShortcutEntries

            $merged = Merge-And-DeduplicateApps -Installed $installed -Shortcuts $shortcuts -Executables @()
            $global:ApplicationInventoryResults = $merged

            Write-Host "" 
            Write-Ok "Quick Scan complete (unique user-facing apps only)."
            Write-Host "  Installed program entries: $($installed.Count)" -ForegroundColor $neutralSilver
            Write-Host "  Shortcut entries:          $($shortcuts.Count)" -ForegroundColor $neutralSilver
            Write-Host "  Unique apps (merged):      $($merged.Count)" -ForegroundColor $neutralSilver

            Write-Host "" 
            Write-Info "Preview of results (top 50 entries)..."
            $merged | Select-Object -First 50 Type, Name, Path, TargetPath, Version, Publisher | Format-Table -AutoSize

            Export-Results -Results $merged
        }
        '2' {
            Write-Warn "Full Scan can take several minutes on some systems."
            $confirm = Read-Host "Continue with Full Scan? (Y/N)"
            if ($confirm -notmatch '^[Yy]') {
                Write-Info "Full Scan cancelled."
                continue
            }

            Write-Info "Starting Full Scan (registry + shortcuts + executables)..."
            $installed   = Get-InstalledPrograms
            $shortcuts   = Get-ShortcutEntries
            $executables = Get-ExecutableEntries

            $merged = Merge-And-DeduplicateApps -Installed $installed -Shortcuts $shortcuts -Executables $executables
            $global:ApplicationInventoryResults = $merged

            Write-Host "" 
            Write-Ok "Full Scan complete (unique user-facing apps only)."
            Write-Host "  Installed program entries: $($installed.Count)" -ForegroundColor $neutralSilver
            Write-Host "  Shortcut entries:          $($shortcuts.Count)" -ForegroundColor $neutralSilver
            Write-Host "  Executable entries:        $($executables.Count)" -ForegroundColor $neutralSilver
            Write-Host "  Unique apps (merged):      $($merged.Count)" -ForegroundColor $neutralSilver

            Write-Host "" 
            Write-Info "Preview of results (top 50 entries)..."
            $merged | Select-Object -First 50 Type, Name, Path, TargetPath, Version, Publisher | Format-Table -AutoSize

            Export-Results -Results $merged
        }
        '3' {
            Write-Info "Exiting Application Inventory Scanner."
            $done = $true
        }
        Default {
            Write-Warn "Invalid selection. Please choose 1, 2, or 3."
        }
    }

    if (-not $done) {
        Write-Host "" 
        $continue = Read-Host "Run another scan? (Y/N)"
        if ($continue -notmatch '^[Yy]') {
            $done = $true
        }
    }
}

# ============================== END OF SCRIPT ==============================
Write-Host "" 
Write-Host "──────────────────────────────────────────────" -ForegroundColor Yellow
Write-Host "  Press Enter to exit..." -ForegroundColor Gray
Write-Host "──────────────────────────────────────────────" -ForegroundColor Yellow
Read-Host
