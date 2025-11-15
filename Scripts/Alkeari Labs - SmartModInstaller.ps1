<#
.SYNOPSIS
    Smart Mod Installer for Bannerlord Environment Manager
.DESCRIPTION
    This script allows users to browse for a working directory containing mod archives,
    a target directory (Modules folder), and then extracts the archives to the target
    using the smart mod installation concept. After extraction, the archives are moved
    to an "Installed" folder within the working directory.
    
    This version works entirely with graphical dialogs and has no command-line interface.
    It caches the last used directories to avoid repetitive navigation.
    
    New: Supports patch-style archives without SubModule.xml.
    - If an archive contains a Modules/<ExistingModuleName>/... structure (or equivalent),
      its contents will be merged into the matching existing module folder(s) in the
      target Modules directory using safe locking and retry logic. If no match is found,
      the installer reports the original "SubModule.xml not found" error.
#>

# No command-line parameters - works entirely with graphical dialogs

# Define config file path
$configFilePath = Join-Path $env:LOCALAPPDATA "BannerlordSmartInstaller\config.xml"

# Try to import shared theming module if present
$themingModule = Join-Path $PSScriptRoot 'Alkeari.AlkTheming.psm1'
if (Test-Path -LiteralPath $themingModule) {
    Import-Module $themingModule -ErrorAction SilentlyContinue
}

# Fallback theme helper if module not present
if (-not (Get-Command -Name New-AlkeariModInstallerTheme -ErrorAction SilentlyContinue)) {
    function New-AlkeariModInstallerTheme {
        [CmdletBinding()]
        param()
        [pscustomobject]@{
            TitleColor = 'Green'
            InfoColor  = 'Cyan'
            WarnColor  = 'Yellow'
            ErrorColor = 'Red'
        }
    }
}

# Function to extract an archive file (supports multiple formats)
function Expand-ArchiveFile
{
    param(
        [string]$ArchivePath,
        [string]$DestinationPath
    )

    try
    {
        # Create destination directory if it doesn't exist
        if (-not (Test-Path $DestinationPath))
        {
            New-Item -ItemType Directory -Path $DestinationPath -Force | Out-Null
        }

        # Get the file extension
        $extension = [System.IO.Path]::GetExtension($ArchivePath).ToLower()

        # Try different extraction methods based on file extension
        switch ($extension)
        {
            ".zip" {
                # Extract ZIP files using built-in PowerShell command
                Expand-Archive -Path $ArchivePath -DestinationPath $DestinationPath -Force
                return $true
            }
            default {
                # For other formats, try PeaZip first, then 7-Zip, then WinRAR
                $archiveFileName = [System.IO.Path]::GetFileName($ArchivePath)

                # Try PeaZip
                $peaZipPaths = @(
                    "${env:ProgramFiles}\PeaZip\peazip.exe",
                    "${env:ProgramFiles}\PeaZip\pea.exe",
                    "${env:ProgramFiles(x86)}\PeaZip\peazip.exe",
                    "${env:ProgramFiles(x86)}\PeaZip\pea.exe"
                )

                foreach ($peaZipPath in $peaZipPaths)
                {
                    if (Test-Path $peaZipPath)
                    {
                        $arguments = "-ext2simple `"$DestinationPath`" `"$ArchivePath`""
                        $process = Start-Process -FilePath $peaZipPath -ArgumentList $arguments -WindowStyle Hidden -PassThru
                        # Wait a short time for the process to start
                        Start-Sleep -Milliseconds 500
                        # Wait for exit with a timeout to prevent hanging
                        $process.WaitForExit(30000)  # 30 second timeout
                        if ($process.ExitCode -eq 0)
                        {
                            return $true
                        }
                        break
                    }
                }

                # Try 7-Zip if PeaZip is not found or fails
                $sevenZipPaths = @(
                    "${env:ProgramFiles}\7-Zip\7z.exe",
                    "${env:ProgramFiles(x86)}\7-Zip\7z.exe",
                    "$env:LOCALAPPDATA\Programs\7-Zip\7z.exe"
                )

                foreach ($sevenZipPath in $sevenZipPaths)
                {
                    if (Test-Path $sevenZipPath)
                    {
                        $arguments = "x `"$ArchivePath`" -o`"$DestinationPath`" -y"
                        $process = Start-Process -FilePath $sevenZipPath -ArgumentList $arguments -WindowStyle Hidden -PassThru
                        # Wait a short time for the process to start
                        Start-Sleep -Milliseconds 500
                        # Wait for exit with a timeout to prevent hanging
                        $process.WaitForExit(30000)  # 30 second timeout
                        if ($process.ExitCode -eq 0)
                        {
                            return $true
                        }
                        break
                    }
                }

                # Try WinRAR if other tools are not found or fail
                $winRarPaths = @(
                    "${env:ProgramFiles}\WinRAR\WinRAR.exe",
                    "${env:ProgramFiles(x86)}\WinRAR\WinRAR.exe",
                    "$env:LOCALAPPDATA\Programs\WinRAR\WinRAR.exe"
                )

                foreach ($winRarPath in $winRarPaths)
                {
                    if (Test-Path $winRarPath)
                    {
                        $arguments = "x -o+ `"$ArchivePath`" `"$DestinationPath\`""
                        $process = Start-Process -FilePath $winRarPath -ArgumentList $arguments -WindowStyle Hidden -PassThru
                        # Wait a short time for the process to start
                        Start-Sleep -Milliseconds 500
                        # Wait for exit with a timeout to prevent hanging
                        $process.WaitForExit(30000)  # 30 second timeout
                        if ($process.ExitCode -eq 0)
                        {
                            return $true
                        }
                        break
                    }
                }

                # If we get here, no extraction tool was found or succeeded
                Write-Error "Failed to extract $ArchivePath : No suitable extraction tool found or extraction failed"
                return $false
            }
        }
    }
    catch
    {
        Write-Error "Failed to extract $ArchivePath : $_"
        return $false
    }
}

# Function to find the module root folder in an extracted directory
function Get-ModuleRootFolder
{
    param(
        [string]$ExtractionPath
    )

    # Look for SubModule.xml to identify the module root
    $subModuleFiles = Get-ChildItem -Path $ExtractionPath -Recurse -Filter "SubModule.xml" -ErrorAction SilentlyContinue

    foreach ($subModuleFile in $subModuleFiles)
    {
        # The folder containing SubModule.xml should be the module root
        $moduleRoot = $subModuleFile.Directory.FullName
        return $moduleRoot
    }

    # If no SubModule.xml found, return nothing
    return $null
}

# Function to show directory browser dialog
function Show-DirectoryBrowserDialog
{
    param(
        [string]$Description = "Select a directory",
        [string]$InitialDirectory = ""
    )

    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = $Description
    $dialog.ShowNewFolderButton = $true

    if ($InitialDirectory -and (Test-Path $InitialDirectory))
    {
        $dialog.SelectedPath = $InitialDirectory
    }

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
    {
        return $dialog.SelectedPath
    }
    return $null
}

# Function to load cached directories
function Get-CachedDirectories
{
    if (Test-Path $configFilePath)
    {
        try
        {
            $config = Import-Clixml -Path $configFilePath
            return $config
        }
        catch
        {
            Write-Host "Error loading cached directories: $_" -ForegroundColor Red
        }
    }
    return @{ WorkingDirectory = $null; TargetDirectory = $null }
}

# Function to save directories to cache
function Save-CachedDirectories
{
    param(
        [string]$WorkingDirectory,
        [string]$TargetDirectory
    )

    try
    {
        # Create directory if it doesn't exist
        $configDir = Split-Path $configFilePath -Parent
        if (-not (Test-Path $configDir))
        {
            New-Item -ItemType Directory -Path $configDir -Force | Out-Null
        }

        # Create config object
        $config = @{
            WorkingDirectory = $WorkingDirectory
            TargetDirectory = $TargetDirectory
            LastUsed = Get-Date
        }

        # Save config
        $config | Export-Clixml -Path $configFilePath -Force
    }
    catch
    {
        Write-Host "Error saving directories to cache: $_" -ForegroundColor Red
    }
}

# Pause helper to keep the window open until user confirms exit
function Pause-BeforeExit
{
    try
    {
        Write-Host "Press Enter to close this window..." -ForegroundColor DarkGray
        Read-Host | Out-Null
    }
    catch
    {
        # Ignore any input errors
    }
}

# Alkeari Labs LLC - Smart Mod Installer (Modernized shell)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Start-AlkeariSmartModInstaller {
    [CmdletBinding()]
    param()

    $theme = New-AlkeariModInstallerTheme

    # Main script execution
    Write-Host "Bannerlord Environment Manager - Smart Mod Installer" -ForegroundColor Green
    Write-Host "=====================================================" -ForegroundColor Green
    Write-Host ""

    # Load cached directories
    $cachedDirs = Get-CachedDirectories

    # Get working directory using dialog
    Write-Host "Please select the working directory containing mod archives..." -ForegroundColor Yellow
    $WorkingDirectory = Show-DirectoryBrowserDialog -Description "Select working directory with mod archives" -InitialDirectory $cachedDirs.WorkingDirectory

    if (-not $WorkingDirectory)
    {
        Write-Host "No working directory selected. Exiting." -ForegroundColor Red
        Write-Host "Script execution finished. Close this window to exit."
        Pause-BeforeExit
        return
    }

    # Validate working directory
    if (-not (Test-Path $WorkingDirectory))
    {
        Write-Host "Working directory does not exist: $WorkingDirectory" -ForegroundColor Red
        Write-Host "Script execution finished. Close this window to exit."
        Pause-BeforeExit
        return
    }

    Write-Host "Working directory: $WorkingDirectory" -ForegroundColor Cyan

    # Get target directory using dialog
    Write-Host "Please select the target Modules directory..." -ForegroundColor Yellow
    $TargetDirectory = Show-DirectoryBrowserDialog -Description "Select target Modules directory" -InitialDirectory $cachedDirs.TargetDirectory

    if (-not $TargetDirectory)
    {
        Write-Host "No target directory selected. Exiting." -ForegroundColor Red
        Write-Host "Script execution finished. Close this window to exit."
        Pause-BeforeExit
        return
    }

    # Validate target directory
    if (-not (Test-Path $TargetDirectory))
    {
        Write-Host "Target directory does not exist: $TargetDirectory" -ForegroundColor Red
        Write-Host "Script execution finished. Close this window to exit."
        Pause-BeforeExit
        return
    }

    Write-Host "Target directory: $TargetDirectory" -ForegroundColor Cyan

    # Find all archive files in working directory
    $zipFiles = Get-ChildItem -Path $WorkingDirectory -Filter "*.zip" -File
    $sevenZFiles = Get-ChildItem -Path $WorkingDirectory -Filter "*.7z" -File
    $rarFiles = Get-ChildItem -Path $WorkingDirectory -Filter "*.rar" -File
    $tarFiles = Get-ChildItem -Path $WorkingDirectory -Filter "*.tar" -File
    $gzFiles = Get-ChildItem -Path $WorkingDirectory -Filter "*.gz" -File
    $bz2Files = Get-ChildItem -Path $WorkingDirectory -Filter "*.bz2" -File
    $arjFiles = Get-ChildItem -Path $WorkingDirectory -Filter "*.arj" -File
    $cabFiles = Get-ChildItem -Path $WorkingDirectory -Filter "*.cab" -File
    $lzhFiles = Get-ChildItem -Path $WorkingDirectory -Filter "*.lzh" -File
    $zFiles = Get-ChildItem -Path $WorkingDirectory -Filter "*.z" -File

    # Combine all archive files into a single array
    $archives = @()
    if ($zipFiles)
    {
        $archives += $zipFiles
    }
    if ($sevenZFiles)
    {
        $archives += $sevenZFiles
    }
    if ($rarFiles)
    {
        $archives += $rarFiles
    }
    if ($tarFiles)
    {
        $archives += $tarFiles
    }
    if ($gzFiles)
    {
        $archives += $gzFiles
    }
    if ($bz2Files)
    {
        $archives += $bz2Files
    }
    if ($arjFiles)
    {
        $archives += $arjFiles
    }
    if ($cabFiles)
    {
        $archives += $cabFiles
    }
    if ($lzhFiles)
    {
        $archives += $lzhFiles
    }
    if ($zFiles)
    {
        $archives += $zFiles
    }

    if ($archives.Count -eq 0)
    {
        Write-Host "No archive files found in the working directory." -ForegroundColor Red
        Write-Host "Script execution finished. Close this window to exit."
        Pause-BeforeExit
        return
    }

    Write-Host "Found $( $archives.Count ) archive file(s) to process." -ForegroundColor Cyan
    Write-Host ""

    # Create Installed directory
    $installedDir = Join-Path $WorkingDirectory "Installed"
    if (-not (Test-Path $installedDir))
    {
        New-Item -ItemType Directory -Path $installedDir | Out-Null
        Write-Host "Created 'Installed' directory: $installedDir" -ForegroundColor Green
    }

    # Process archive files with dynamic slot management
    $maxConcurrent = 5
    $jobQueue = [System.Collections.ArrayList]@()
    $processedCount = 0
    $totalCount = $archives.Count
    $successCount = 0
    $failCount = 0
    $failedItems = @()

    # Initialize the job queue with all archives
    foreach ($archive in $archives)
    {
        $jobQueue.Add($archive) | Out-Null
    }

    Write-Host "Processing $totalCount archive(s) with up to $maxConcurrent concurrent extractions..." -ForegroundColor Cyan

    # Active jobs tracking
    $activeJobs = @{ }

    # Process archives dynamically - start new ones as slots become free
    while ($jobQueue.Count -gt 0 -or $activeJobs.Count -gt 0)
    {
        # Start new jobs if we have slots available
        while ($activeJobs.Count -lt $maxConcurrent -and $jobQueue.Count -gt 0)
        {
            $archive = $jobQueue[0]
            $jobQueue.RemoveAt(0)

            Write-Host "  Starting: $( $archive.Name )" -ForegroundColor Yellow

            # Start a background job for the archive
            $job = Start-Job -ScriptBlock {
                param($archivePath, $targetDirectory, $installedDir, $archiveName)

                $ErrorActionPreference = 'Stop'

                # Create a temporary extraction directory
                $tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
                New-Item -ItemType Directory -Path $tempDir -Force -ErrorAction Stop | Out-Null

                try
                {
                    # Define functions needed in the job context
                    function Get-ModuleRootFolder
                    {
                        param([string]$ExtractionPath)

                        $subModuleFiles = Get-ChildItem -Path $ExtractionPath -Recurse -Filter "SubModule.xml" -ErrorAction SilentlyContinue

                        foreach ($subModuleFile in $subModuleFiles)
                        {
                            $moduleRoot = $subModuleFile.Directory.FullName
                            return $moduleRoot
                        }

                        return $null
                    }

                    # Try to extract with built-in ZIP support first
                    $extension = [System.IO.Path]::GetExtension($archivePath).ToLower()
                    $extracted = $false

                    if ($extension -eq ".zip")
                    {
                        try
                        {
                            Expand-Archive -Path $archivePath -DestinationPath $tempDir -Force -ErrorAction Stop
                            $extracted = $true
                        }
                        catch
                        {
                            # Fall back to other tools if built-in fails
                        }
                    }

                    # If not ZIP or ZIP extraction failed, try external tools
                    if (-not $extracted)
                    {
                        # Try PeaZip
                        $peaZipPaths = @(
                            "${env:ProgramFiles}\PeaZip\peazip.exe",
                            "${env:ProgramFiles}\PeaZip\pea.exe",
                            "${env:ProgramFiles(x86)}\PeaZip\peazip.exe",
                            "${env:ProgramFiles(x86)}\PeaZip\pea.exe"
                        )

                        foreach ($peaZipPath in $peaZipPaths)
                        {
                            if (Test-Path $peaZipPath)
                            {
                                $arguments = "-ext2simple `"$tempDir`" `"$archivePath`""
                                $process = Start-Process -FilePath $peaZipPath -ArgumentList $arguments -WindowStyle Hidden -PassThru
                                Start-Sleep -Milliseconds 500
                                $process.WaitForExit(30000)
                                if ($process.ExitCode -eq 0)
                                {
                                    $extracted = $true
                                    break
                                }
                            }
                        }

                        # Try 7-Zip if PeaZip failed
                        if (-not $extracted)
                        {
                            $sevenZipPaths = @(
                                "${env:ProgramFiles}\7-Zip\7z.exe",
                                "${env:ProgramFiles(x86)}\7-Zip\7z.exe",
                                "$env:LOCALAPPDATA\Programs\7-Zip\7z.exe"
                            )

                            foreach ($sevenZipPath in $sevenZipPaths)
                            {
                                if (Test-Path $sevenZipPath)
                                {
                                    $arguments = "x `"$archivePath`" -o`"$tempDir`" -y"
                                    $process = Start-Process -FilePath $sevenZipPath -ArgumentList $arguments -WindowStyle Hidden -PassThru
                                    Start-Sleep -Milliseconds 500
                                    $process.WaitForExit(30000)
                                    if ($process.ExitCode -eq 0)
                                    {
                                        $extracted = $true
                                        break
                                    }
                                }
                            }
                        }

                        # Try WinRAR if 7-Zip failed
                        if (-not $extracted)
                        {
                            $winRarPaths = @(
                                "${env:ProgramFiles}\WinRAR\WinRAR.exe",
                                "${env:ProgramFiles(x86)}\WinRAR\WinRAR.exe",
                                "$env:LOCALAPPDATA\Programs\WinRAR\WinRAR.exe"
                            )

                            foreach ($winRarPath in $winRarPaths)
                            {
                                if (Test-Path $winRarPath)
                                {
                                    $arguments = "x -o+ `"$archivePath`" `"$tempDir\`""
                                    $process = Start-Process -FilePath $winRarPath -ArgumentList $arguments -WindowStyle Hidden -PassThru
                                    Start-Sleep -Milliseconds 500
                                    $process.WaitForExit(30000)
                                    if ($process.ExitCode -eq 0)
                                    {
                                        $extracted = $true
                                        break
                                    }
                                }
                            }
                        }
                    }

                    if (-not $extracted)
                    {
                        return @{ Success = $false; Message = "Failed to extract ${archiveName}"; ArchiveName = $archiveName }
                    }

                    # Find the module root folder
                    $moduleRoots = @(Get-ModuleRootFolder -ExtractionPath $tempDir)

                    # If no SubModule.xml roots were found, fall back to multi-module patch handling
                    if (-not $moduleRoots -or $moduleRoots.Count -eq 0)
                    {
                        # Fallback path: support patch-style archives without SubModule.xml
                        # These archives typically contain one or more Modules/<ExistingModuleName>/... structures
                        try
                        {
                            $patched = New-Object System.Collections.Generic.List[string]

                            function Invoke-WithRetryLocal
                            {
                                param([ScriptBlock]$Action, [int]$RetryCount = 5, [int]$DelayMs = 300)
                                for ($i = 1; $i -le $RetryCount; $i++) {
                                    try
                                    {
                                        & $Action; return $true
                                    }
                                    catch
                                    {
                                        if ($i -eq $RetryCount)
                                        {
                                            throw
                                        }
                                        Start-Sleep -Milliseconds $DelayMs
                                        $DelayMs = [Math]::Min(4000, [int]([double]$DelayMs * 1.75))
                                    }
                                }
                                return $false
                            }

                            function Clear-ReadOnlyAttributesLocal
                            {
                                param([string]$Path)
                                if (Test-Path $Path)
                                {
                                    try
                                    {
                                        (Get-Item -LiteralPath $Path -Force).Attributes = 'Directory'
                                    }
                                    catch
                                    {
                                    }
                                    Get-ChildItem -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue |
                                            ForEach-Object {
                                                try
                                                {
                                                    $_.Attributes = 'Normal'
                                                }
                                                catch
                                                {
                                                }
                                            }
                                }
                            }

                            $candidatesRoot = Join-Path $tempDir 'Modules'
                            $candidateDirs = @()
                            if (Test-Path $candidatesRoot)
                            {
                                $candidateDirs = Get-ChildItem -Path $candidatesRoot -Directory -ErrorAction SilentlyContinue
                            }
                            if (-not $candidateDirs -or $candidateDirs.Count -eq 0)
                            {
                                # Search for wrapped archives that contain a nested 'Modules' directory
                                $modulesDirs = Get-ChildItem -Path $tempDir -Recurse -Directory -Filter 'Modules' -ErrorAction SilentlyContinue
                                foreach ($mDir in $modulesDirs)
                                {
                                    try
                                    {
                                        $children = Get-ChildItem -Path $mDir.FullName -Directory -ErrorAction SilentlyContinue
                                        if ($children)
                                        {
                                            $candidateDirs += $children
                                        }
                                    }
                                    catch
                                    {
                                    }
                                }
                            }
                            if (-not $candidateDirs -or $candidateDirs.Count -eq 0)
                            {
                                # As a last resort, consider top-level directories under temp as potential targets
                                $candidateDirs = Get-ChildItem -Path $tempDir -Directory -ErrorAction SilentlyContinue
                            }

                            $existingModules = @{ }
                            Get-ChildItem -Path $targetDirectory -Directory -ErrorAction SilentlyContinue | ForEach-Object {
                                $existingModules[$_.Name.ToLower()] = $_.FullName
                            }

                            foreach ($cand in $candidateDirs)
                            {
                                $name = $cand.Name
                                $key = $name.ToLower()
                                if ( $existingModules.ContainsKey($key))
                                {
                                    $destDir = $existingModules[$key]
                                    # Per-module mutex to avoid concurrent writes
                                    $mutexName = 'Global\BEM_SMI_' + ($name -replace '\s+', '_')
                                    $mutex = New-Object System.Threading.Mutex($false, $mutexName)
                                    $got = $false
                                    try
                                    {
                                        $got = $mutex.WaitOne([TimeSpan]::FromMinutes(2))
                                        if (-not $got)
                                        {
                                            continue
                                        }
                                        if (-not (Test-Path $destDir))
                                        {
                                            New-Item -ItemType Directory -Path $destDir -Force -ErrorAction Stop | Out-Null
                                        }
                                        Clear-ReadOnlyAttributesLocal -Path $destDir
                                        Invoke-WithRetryLocal -Action { Copy-Item -Path (Join-Path $cand.FullName '*') -Destination $destDir -Recurse -Force -ErrorAction Stop } | Out-Null
                                        $patched.Add($name) | Out-Null
                                    }
                                    finally
                                    {
                                        if ($got)
                                        {
                                            try
                                            {
                                                $mutex.ReleaseMutex() | Out-Null
                                            }
                                            catch
                                            {
                                            }
                                        }
                                        $mutex.Dispose()
                                    }
                                }
                            }

                            if ($patched.Count -gt 0)
                            {
                                $destinationArchive = Join-Path $installedDir $archiveName
                                Invoke-WithRetryLocal -Action { Move-Item -Path $archivePath -Destination $destinationArchive -Force -ErrorAction Stop } | Out-Null
                                $patchedList = [string]::Join(', ', $patched)
                                return @{ Success = $true; Message = "Installed patch into existing module(s): $patchedList"; ArchiveName = $archiveName }
                            }
                        }
                        catch
                        {
                            $err = $_.Exception.Message
                            return @{ Success = $false; Message = "Error processing ${archiveName}: ${err}"; ArchiveName = $archiveName }
                        }

                        return @{ Success = $false; Message = "Could not find SubModule.xml in the extracted files"; ArchiveName = $archiveName }
                    }

                    # At least one module root found (one or more SubModule.xml). Install each root as its own module
                    $installedModules = New-Object System.Collections.Generic.List[string]

                    foreach ($moduleRoot in $moduleRoots)
                    {
                        # Get module name from SubModule.xml (using Name instead of Id)
                        [xml]$subModuleXml = Get-Content (Join-Path $moduleRoot "SubModule.xml")
                        $moduleName = $subModuleXml.Module.Name.Value
                        if (-not $moduleName)
                        {
                            $moduleName = $subModuleXml.Module.Id.Value
                        }
                        if (-not $moduleName)
                        {
                            $moduleName = [System.IO.Path]::GetFileNameWithoutExtension($archiveName)
                        }

                        # Special case for Harmony mod - must be named "Bannerlord.Harmony" for BLSE compatibility
                        if ($moduleName -eq "Harmony" -or $moduleName -eq "Bannerlord.Harmony")
                        {
                            $moduleName = "Bannerlord.Harmony"
                        }

                        $moduleNameOriginal = $moduleName

                        # Sanitize module name for filesystem folder usage
                        function Sanitize-ModuleFolderName
                        {
                            param([string]$Name)
                            if (-not $Name)
                            {
                                return 'Module'
                            }
                            $invalid = [Regex]::Escape(-join ([System.IO.Path]::GetInvalidFileNameChars()))
                            $san = [Regex]::Replace($Name, "[${invalid}]", '-')
                            $san = ($san -replace '\s+', ' ').Trim()
                            # Remove trailing periods and spaces which are invalid for folder names
                            $san = $san.TrimEnd('.', ' ')
                            if ( [string]::IsNullOrWhiteSpace($san))
                            {
                                $san = 'Module'
                            }
                            return $san
                        }

                        $moduleFolderName = Sanitize-ModuleFolderName -Name $moduleName
                        $targetModuleDir = Join-Path $targetDirectory $moduleFolderName

                        function Invoke-WithRetry
                        {
                            param([ScriptBlock]$Action, [int]$RetryCount = 5, [int]$DelayMs = 300)
                            for ($i = 1; $i -le $RetryCount; $i++) {
                                try
                                {
                                    & $Action; return
                                }
                                catch
                                {
                                    if ($i -eq $RetryCount)
                                    {
                                        throw
                                    }
                                    Start-Sleep -Milliseconds $DelayMs
                                    $DelayMs = [Math]::Min(4000, [int]([double]$DelayMs * 1.75))
                                }
                            }
                        }

                        function Clear-ReadOnlyAttributes
                        {
                            param([string]$Path)
                            if (Test-Path $Path)
                            {
                                try
                                {
                                    (Get-Item -LiteralPath $Path -Force).Attributes = 'Directory'
                                }
                                catch
                                {
                                }
                                Get-ChildItem -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue |
                                        ForEach-Object {
                                            try
                                            {
                                                $_.Attributes = 'Normal'
                                            }
                                            catch
                                            {
                                            }
                                        }
                            }
                        }

                        # Per-module cross-process mutex to avoid concurrent writes
                        $mutexName = 'Global\BEM_SMI_' + ($moduleFolderName -replace '\s+', '_')
                        $mutex = New-Object System.Threading.Mutex($false, $mutexName)
                        $acquired = $false
                        try
                        {
                            $acquired = $mutex.WaitOne([TimeSpan]::FromMinutes(2))
                            if (-not $acquired)
                            {
                                return @{ Success = $false; Message = "Timeout waiting for module lock: ${moduleNameOriginal}"; ArchiveName = $archiveName }
                            }

                            # Remove existing module directory if it exists (with retries and attribute clearing)
                            if (Test-Path $targetModuleDir)
                            {
                                Clear-ReadOnlyAttributes -Path $targetModuleDir
                                Invoke-WithRetry -Action { Remove-Item -LiteralPath $targetModuleDir -Recurse -Force -ErrorAction Stop }
                            }

                            # Create target directory if it doesn't exist
                            if (-not (Test-Path $targetModuleDir))
                            {
                                Invoke-WithRetry -Action { New-Item -ItemType Directory -Path $targetModuleDir -Force -ErrorAction Stop | Out-Null }
                            }

                            # Copy all files from this module root to target directory
                            Invoke-WithRetry -Action { Copy-Item -Path "$moduleRoot\*" -Destination $targetModuleDir -Recurse -Force -ErrorAction Stop }

                            $installedModules.Add($moduleNameOriginal) | Out-Null
                        }
                        catch
                        {
                            $errorMessage = $_.Exception.Message
                            if ($_.Exception -is [System.UnauthorizedAccessException])
                            {
                                $errorMessage = $errorMessage + " Hint: Close the game/launchers and tools that may lock files, then retry."
                            }
                            return @{ Success = $false; Message = "Error processing ${archiveName}: ${errorMessage}"; ArchiveName = $archiveName }
                        }
                        finally
                        {
                            if ($acquired)
                            {
                                try
                                {
                                    $mutex.ReleaseMutex() | Out-Null
                                }
                                catch
                                {
                                }
                            }
                            $mutex.Dispose()
                        }
                    }

                    # Move archive file to Installed directory once after all module roots have been processed
                    $destinationArchive = Join-Path $installedDir $archiveName
                    try
                    {
                        Move-Item -Path $archivePath -Destination $destinationArchive -Force -ErrorAction Stop
                    }
                    catch
                    {
                        # Non-fatal: log via message, but continue
                        return @{ Success = $false; Message = "Installed modules (${($installedModules -join ', ')}) but failed moving archive: ${archiveName}. Error: $($_.Exception.Message)"; ArchiveName = $archiveName }
                    }

                    if ($installedModules.Count -gt 0)
                    {
                        $names = $installedModules -join ', '
                        return @{ Success = $true; Message = "Successfully installed module(s): ${names}"; ArchiveName = $archiveName }
                    }
                    else
                    {
                        return @{ Success = $false; Message = "No module roots were installed for ${archiveName}"; ArchiveName = $archiveName }
                    }
                }
                catch
                {
                    # Properly handle the exception message using the ${} syntax
                    $errorMessage = $_.Exception.Message
                    return @{ Success = $false; Message = "Error processing ${archiveName}: ${errorMessage}"; ArchiveName = $archiveName }
                }
                finally
                {
                    # Clean up temporary directory
                    if (Test-Path $tempDir)
                    {
                        Remove-Item -Path $tempDir -Recurse -Force
                    }
                }
            } -ArgumentList $archive.FullName, $TargetDirectory, $installedDir, $archive.Name

            # Add job to active jobs tracking
            $activeJobs[$job.Name] = @{
                Job = $job
                ArchiveName = $archive.Name
            }
        }

        # Check for completed jobs if we have active jobs
        if ($activeJobs.Count -gt 0)
        {
            # Get completed jobs
            $completedJobs = $activeJobs.GetEnumerator() | Where-Object { $_.Value.Job.State -eq "Completed" }

            # Process results of completed jobs
            foreach ($completedJob in $completedJobs)
            {
                $jobInfo = $completedJob.Value
                $job = $jobInfo.Job
                $archiveName = $jobInfo.ArchiveName

                try
                {
                    $result = Receive-Job -Job $job
                    if ($result.Success)
                    {
                        Write-Host "  Completed: $( $result.Message )" -ForegroundColor Green
                        $successCount++
                    }
                    else
                    {
                        Write-Host "  Failed: $( $result.Message )" -ForegroundColor Red
                        $failCount++
                        $failedItems += ('{0}: {1}' -f $archiveName, $result.Message)
                    }
                }
                catch
                {
                    Write-Host "  Error retrieving result for ${archiveName}: $_" -ForegroundColor Red
                }
                finally
                {
                    # Clean up job
                    Remove-Job -Job $job
                    # Remove from active jobs
                    $activeJobs.Remove($completedJob.Key)
                    # Increment processed count
                    $processedCount++
                    Write-Host "  Progress: $processedCount of $totalCount archives processed" -ForegroundColor Gray
                }
            }

            # Small sleep to prevent excessive CPU usage
            if ($activeJobs.Count -gt 0)
            {
                Start-Sleep -Milliseconds 500
            }
        }
    }

    if ($failCount -gt 0)
    {
        Write-Host "All archives processed (with errors)." -ForegroundColor Yellow
    }
    else
    {
        Write-Host "All archives processed!" -ForegroundColor Green
    }

    # Save directories to cache
    Save-CachedDirectories -WorkingDirectory $WorkingDirectory -TargetDirectory $TargetDirectory

    # Summary
    Write-Host ("Summary: {0} succeeded, {1} failed." -f $successCount, $failCount) -ForegroundColor Cyan
    if ($failCount -gt 0)
    {
        Write-Host "One or more archives failed to install:" -ForegroundColor Yellow
        foreach ($item in $failedItems)
        {
            Write-Host ("  - {0}" -f $item) -ForegroundColor Red
        }
        Write-Host "Installation process completed with errors." -ForegroundColor Yellow
    }
    else
    {
        Write-Host "Installation process completed successfully." -ForegroundColor Green
    }

    Write-Host "Installed modules can be found in: $TargetDirectory" -ForegroundColor Cyan
    if ($failCount -gt 0)
    {
        Write-Host "Successfully processed archives have been moved to: $installedDir" -ForegroundColor Cyan
        Write-Host "Some failed archives may remain in the working directory: $WorkingDirectory" -ForegroundColor Yellow
    }
    else
    {
        Write-Host "Processed archives have been moved to: $installedDir" -ForegroundColor Cyan
    }

    Write-Host "Directories have been cached for next use." -ForegroundColor Green
    Write-Host ""
    Write-Host "Script execution finished. Close this window to exit."
    Pause-BeforeExit
}

if ($MyInvocation.InvocationName -ne '.') {
    Start-AlkeariSmartModInstaller
}
