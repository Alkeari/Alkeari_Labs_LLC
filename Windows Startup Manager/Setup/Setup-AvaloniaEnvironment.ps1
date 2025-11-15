<# ======================================================================
  Alkeari Labs LLC — Avalonia project setup (Windows, Linux, macOS)
  Action:
   1) Backup csproj
   2) Remove all PackageReference entries
   3) Install curated non-deprecated packages (pinned)
   4) Ensure multi-runtime publish properties
   5) Mark Diagnostics as PrivateAssets=all
====================================================================== #>

# ---------- Config (edit if your paths change) ----------
$SolutionDir = 'C:\Users\Josep\RiderProjects\Alkeari Labs LLC'
$ProjectName = 'Windows Startup Manager'
$ProjectDir = Join-Path $SolutionDir $ProjectName
$CsprojPath = Join-Path $ProjectDir  'Windows Startup Manager.csproj'

# ---------- UI / Theming ----------
# Alkeari Labs LLC palette: Black #000000, Gold #DAA520, Brown #8B4513, Silver #C0C0C0, Soft White #F5F5F5
Clear-Host
$banner = @"
======================================================================
                        Alkeari Labs LLC
                 Avalonia Desktop Setup Assistant
======================================================================
Solution: $SolutionDir
Project : $ProjectName
CSProj  : $CsprojPath
======================================================================
"@
Write-Host $banner -ForegroundColor Yellow

# ---------- Pre-flight checks ----------
if (-not (Get-Command dotnet -ErrorAction SilentlyContinue))
{
    Write-Host "Error: .NET SDK not found on PATH. Please install .NET 6 SDK or newer." -ForegroundColor Red
    exit
}
if (-not (Test-Path -LiteralPath $CsprojPath))
{
    Write-Host "Error: Project file not found at:" -ForegroundColor Red
    Write-Host "       $CsprojPath" -ForegroundColor Red
    exit
}

# ---------- Helper: safe exec ----------
function Invoke-Cli
{
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [Parameter()][string[]]$Parameters
    )
    # Properly quote parameters that might contain spaces
    $quotedParams = $Parameters | ForEach-Object {
        if ($_ -match '\s' -and $_ -notmatch '^".*"$')
        {
            '"' + $_ + '"'
        }
        else
        {
            $_
        }
    }
    Write-Host "> $FilePath $( $quotedParams -join ' ' )" -ForegroundColor Gray
    $p = Start-Process -FilePath $FilePath -ArgumentList $quotedParams -NoNewWindow -PassThru -Wait
    if ($p.ExitCode -ne 0)
    {
        throw "Command failed with exit code $( $p.ExitCode ): $FilePath $( $quotedParams -join ' ' )"
    }
}

# ---------- 1) Backup csproj ----------
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$BackupPath = "$CsprojPath.$stamp.bak"
Copy-Item -LiteralPath $CsprojPath -Destination $BackupPath -Force
Write-Host "Backed up csproj to: $BackupPath" -ForegroundColor Green

# ---------- 2) Remove all packages ----------
[xml]$xml = Get-Content -LiteralPath $CsprojPath
$ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
$ns.AddNamespace('msb', $xml.Project.NamespaceURI)

# Collect existing PackageReference names
$pkgNodes = $xml.SelectNodes('//msb:PackageReference', $ns)
$existingPkgs = @()
foreach ($n in $pkgNodes)
{
    $name = $n.GetAttribute('Include')
    if ($name)
    {
        $existingPkgs += $name
    }
}

if ($existingPkgs.Count -gt 0)
{
    Write-Host "Removing existing packages from project..." -ForegroundColor Yellow
    foreach ($name in $existingPkgs | Select-Object -Unique)
    {
        try
        {
            Invoke-Cli -FilePath 'dotnet' -Parameters @('remove', $CsprojPath, 'package', $name)
        }
        catch
        {
            Write-Host "Warning: failed to remove $name. Will continue." -ForegroundColor DarkYellow
        }
    }
}
else
{
    Write-Host "No existing PackageReference entries found." -ForegroundColor Yellow
}

# ---------- 3) Install curated packages (pinned, non-deprecated) ----------
$packages = @(
    @{ Name = 'Avalonia'; Version = '11.3.8' }
    @{ Name = 'Avalonia.Desktop'; Version = '11.3.8' }
    @{ Name = 'Avalonia.Skia'; Version = '11.3.8' }
    @{ Name = 'Avalonia.Themes.Fluent'; Version = '11.3.8' }
    @{ Name = 'Avalonia.Fonts.Inter'; Version = '11.3.8' }

    @{ Name = 'Avalonia.Controls.DataGrid'; Version = '11.3.8' }
    @{ Name = 'Avalonia.Controls.ColorPicker'; Version = '11.3.8' }
    @{ Name = 'Avalonia.AvaloniaEdit'; Version = '11.3.0' }
    @{ Name = 'Avalonia.Markup.Xaml.Loader'; Version = '11.3.8' }

# Linux/FreeDesktop integration
    @{ Name = 'Avalonia.FreeDesktop'; Version = '11.3.8' }

# Dev-time diagnostics and remote protocol
    @{ Name = 'Avalonia.Diagnostics'; Version = '11.3.8' }
    @{ Name = 'Avalonia.Remote.Protocol'; Version = '11.3.8' }

# Optional: headless/test support
    @{ Name = 'Avalonia.Headless'; Version = '11.3.8' }
    @{ Name = 'Avalonia.Headless.XUnit'; Version = '11.3.8' }

# Optional: Windows ANGLE backend (safe to include on Windows)
    @{ Name = 'Avalonia.Angle.Windows.Natives'; Version = '2.1.25547.20250602' }
)

Write-Host "Installing curated Avalonia packages..." -ForegroundColor Yellow
foreach ($p in $packages)
{
    try
    {
        Invoke-Cli -FilePath 'dotnet' -Parameters @('add', $CsprojPath, 'package', $p.Name, '--version', $p.Version)
    }
    catch
    {
        Write-Host "Error installing $( $p.Name ) $( $p.Version ): $( $_.Exception.Message )" -ForegroundColor Red
        throw
    }
}

# ---------- 4) Ensure multi-runtime publish properties & TFM ----------
[xml]$xml = Get-Content -LiteralPath $CsprojPath
$ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
$ns.AddNamespace('msb', $xml.Project.NamespaceURI)

# Ensure (or create) a PropertyGroup
$pg = $xml.SelectSingleNode('/msb:Project/msb:PropertyGroup', $ns)
if (-not $pg)
{
    $pg = $xml.CreateElement('PropertyGroup', $xml.Project.NamespaceURI)
    $xml.Project.AppendChild($pg) | Out-Null
}

function Set-Element
{
    param([xml]$doc, [System.Xml.XmlElement]$group, [string]$name, [string]$value)
    $node = $group.SelectSingleNode("msb:$name", $ns)
    if (-not $node)
    {
        $node = $doc.CreateElement($name, $doc.Project.NamespaceURI); $group.AppendChild($node) | Out-Null
    }
    $node.InnerText = $value
}

# Target .NET 6 for broad runtime compatibility
Set-Element -doc $xml -group $pg -name 'TargetFramework' -value 'net6.0'
# Single-file, self-contained, multi-runtime
Set-Element -doc $xml -group $pg -name 'PublishSingleFile'     -value 'true'
Set-Element -doc $xml -group $pg -name 'SelfContained'          -value 'true'
# Include common desktop RIDs. Add arm64 if you need it later.
Set-Element -doc $xml -group $pg -name 'RuntimeIdentifiers'     -value 'win-x64;linux-x64;osx-x64'

# ---------- 5) Mark Diagnostics as PrivateAssets=all ----------
# Add or update the PackageReference node for Avalonia.Diagnostics
$itemGroup = $xml.SelectSingleNode('/msb:Project/msb:ItemGroup[msb:PackageReference]', $ns)
if (-not $itemGroup)
{
    $itemGroup = $xml.CreateElement('ItemGroup', $xml.Project.NamespaceURI)
    $xml.Project.AppendChild($itemGroup) | Out-Null
}

$diagNode = $xml.SelectSingleNode("//msb:PackageReference[@Include='Avalonia.Diagnostics']", $ns)
if (-not $diagNode)
{
    $diagNode = $xml.CreateElement('PackageReference', $xml.Project.NamespaceURI)
    $diagNode.SetAttribute('Include', 'Avalonia.Diagnostics')
    $diagNode.SetAttribute('Version', '11.3.8')
    $itemGroup.AppendChild($diagNode) | Out-Null
}
$diagNode.SetAttribute('PrivateAssets', 'all')

# Save csproj changes
$xml.Save($CsprojPath)
Write-Host "Updated project file for multi-runtime publish and diagnostics privacy." -ForegroundColor Green

# ---------- Restore ----------
Invoke-Cli -FilePath 'dotnet' -Parameters @('restore', $CsprojPath)

Write-Host ""
Write-Host "All done. You can now publish self-contained, single-file builds for:" -ForegroundColor Green
Write-Host " - win-x64" -ForegroundColor Green
Write-Host " - linux-x64" -ForegroundColor Green
Write-Host " - osx-x64" -ForegroundColor Green
Write-Host ""
Write-Host "Examples:" -ForegroundColor Yellow
Write-Host "  dotnet publish `"$CsprojPath`" -c Release -r win-x64   /p:PublishSingleFile=true /p:SelfContained=true" -ForegroundColor Gray
Write-Host "  dotnet publish `"$CsprojPath`" -c Release -r linux-x64 /p:PublishSingleFile=true /p:SelfContained=true" -ForegroundColor Gray
Write-Host "  dotnet publish `"$CsprojPath`" -c Release -r osx-x64   /p:PublishSingleFile=true /p:SelfContained=true" -ForegroundColor Gray

# ============================== END OF SCRIPT ==============================
Write-Host ""
Write-Host "──────────────────────────────────────────────" -ForegroundColor Yellow
Write-Host "  Press Enter to exit..." -ForegroundColor Gray
Write-Host "──────────────────────────────────────────────" -ForegroundColor Yellow
Read-Host