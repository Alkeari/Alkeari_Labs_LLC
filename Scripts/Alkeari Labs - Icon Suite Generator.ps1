<#
  Alkeari Labs LLC - Icon Suite Generator (Modernized)
  - Self-bootstrapping dependencies in %APPDATA%
  - Multi-file picker via GUI
  - Multi-size ICO generation with Magick.NET / Skia / WIC fallback
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Try to import shared theming module if present
$themingModule = Join-Path $PSScriptRoot 'Alkeari.AlkTheming.psm1'
if (Test-Path -LiteralPath $themingModule) {
    Import-Module $themingModule -ErrorAction SilentlyContinue
}

# If shared module not loaded, fall back to local branding function
if (-not (Get-Command -Name New-AlkeariIconBranding -ErrorAction SilentlyContinue)) {
    function New-AlkeariIconBranding {
        [CmdletBinding()]
        param()
        [pscustomobject]@{
            Gold   = 'Yellow'
            Silver = 'Gray'
            White  = 'White'
            Red    = 'Red'
            Green  = 'Green'
            Cyan   = 'Cyan'
        }
    }
}

# ============================== ROBUST SCRIPT BASE ==============================
# This stays for logging and any local assets; deps live in %APPDATA%
$ScriptBase = $PSScriptRoot
if ( [string]::IsNullOrWhiteSpace($ScriptBase))
{
    $ScriptBase = Split-Path -Parent $MyInvocation.MyCommand.Path 2> $null
    if ( [string]::IsNullOrWhiteSpace($ScriptBase))
    {
        $ScriptBase = (Get-Location).Path
    }
}

# ============================== CONFIG: APPDATA LOCATION ==============================
# Place dependencies under %APPDATA%\Alkeari Labs LLC\IconSuite\deps
$CompanyRoot = Join-Path $env:APPDATA "Alkeari Labs LLC"
$AppRoot = Join-Path $CompanyRoot "IconSuite"
$DepsRoot = Join-Path $AppRoot "deps"
$LoadFolder = Join-Path $DepsRoot "load"

# Pin versions for reproducibility
$Packages = @(
    @{ Id = "Magick.NET-Q8-AnyCPU"; Ver = "13.8.0"; DllPatterns = @("Magick.NET-Q8-AnyCPU*.dll", "Magick.NET*.dll") },
    @{ Id = "SkiaSharp"; Ver = "2.88.6"; DllPatterns = @("SkiaSharp.dll") },
    @{ Id = "SkiaSharp.Views"; Ver = "2.88.6"; DllPatterns = @("SkiaSharp.Views.dll") },
    @{ Id = "SkiaSharp.NativeAssets.Win32"; Ver = "2.88.6"; DllPatterns = @("libSkiaSharp.dll") }, # native assets in runtimes\win-*\native
    @{ Id = "Svg.Skia"; Ver = "0.5.0"; DllPatterns = @("Svg.Skia.dll") }
)

$IconSizes = @(256, 128, 64, 32, 16)

# ============================== DEP BOOTSTRAP (NuGet v3) ==============================
Add-Type -AssemblyName System.IO.Compression.FileSystem
try
{
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
}
catch
{
}

function Get-NuGetV3Url
{
    param([string]$Id, [string]$Version)
    $lid = $Id.ToLowerInvariant()
    $ver = $Version.ToLowerInvariant()
    "https://api.nuget.org/v3-flatcontainer/$lid/$ver/$lid.$ver.nupkg"
}

function Initialize-Dependencies
{
    foreach ($dir in @($CompanyRoot, $AppRoot, $DepsRoot, $LoadFolder))
    {
        if (-not (Test-Path -LiteralPath $dir))
        {
            New-Item -ItemType Directory -Path $dir | Out-Null
        }
    }

    foreach ($pkg in $Packages)
    {
        $id = $pkg.Id
        $ver = $pkg.Ver
        $pkgFolder = Join-Path $DepsRoot "$id.$ver"
        $nupkgPath = Join-Path $DepsRoot "$( $id ).$( $ver ).nupkg"

        if (-not (Test-Path -LiteralPath $pkgFolder))
        {
            Info "Fetching $id $ver..."
            $url = Get-NuGetV3Url -Id $id -Version $ver
            try
            {
                Invoke-WebRequest -UseBasicParsing -Uri $url -OutFile $nupkgPath | Out-Null
            }
            catch
            {
                Err "Download failed: $id $ver - $( $_.Exception.Message )"
                throw
            }

            try
            {
                [System.IO.Compression.ZipFile]::ExtractToDirectory($nupkgPath, $pkgFolder)
            }
            finally
            {
                if (Test-Path -LiteralPath $nupkgPath)
                {
                    Remove-Item $nupkgPath -Force
                }
            }
            Good "Downloaded $id"
        }
        else
        {
            Info "$id $ver already present."
        }

        # Copy managed DLLs under lib/* and native assets we might need
        $managed = Get-ChildItem -LiteralPath $pkgFolder -Recurse -File -ErrorAction SilentlyContinue |
                Where-Object { $_.Extension -ieq ".dll" -and $_.FullName -match "\\lib\\" }

        $native = Get-ChildItem -LiteralPath $pkgFolder -Recurse -File -ErrorAction SilentlyContinue |
                Where-Object { ($pkg.DllPatterns | Where-Object { $PSItem.Name -like $_ }).Count -gt 0 -or $PSItem.FullName -match "\\runtimes\\win-" }

        foreach ($dll in ($managed + $native | Sort-Object FullName -Unique))
        {
            Copy-Item -Path $dll.FullName -Destination (Join-Path $LoadFolder $dll.Name) -Force
        }
    }

    # Add load folder to PATH for native resolution
    $env:PATH = "$LoadFolder;$env:PATH"

    # Load only managed assemblies; native .dlls are resolved by PATH
    foreach ($dll in Get-ChildItem -LiteralPath $LoadFolder -Filter *.dll -File)
    {
        try
        {
            [Reflection.Assembly]::LoadFrom($dll.FullName) | Out-Null
        }
        catch
        {
        }
    }

    # Explicitly load Magick.NET managed assembly and validate availability
    $magickDll = Get-ChildItem -LiteralPath $LoadFolder -Filter "Magick.NET-Q8-AnyCPU*.dll" -File -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($magickDll)
    {
        try
        {
            [Reflection.Assembly]::LoadFrom($magickDll.FullName) | Out-Null
        }
        catch
        {
            Err "Failed to load $( $magickDll.Name ) - $( $_.Exception.Message )"
        }
    }
    else
    {
        Warn "Magick.NET DLL not found in $LoadFolder. Raster work will rely on WIC fallback if needed."
    }

    if (-not ("ImageMagick.MagickImage" -as [type]))
    {
        Warn "Magick.NET not active. Raster conversions may use WIC fallback."
    }
    else
    {
        Good "Magick.NET ready."
    }

    Good "Dependencies ready in $LoadFolder"
}

# ============================== IMAGE RASTER HELPERS ==============================
# Save PNG bytes to file (utility if you want temp debugging)
function Save-PngBytes
{
    param([byte[]]$Bytes, [string]$Dest)
    try
    {
        $dir = [IO.Path]::GetDirectoryName($Dest)
        if (-not (Test-Path -LiteralPath $dir))
        {
            New-Item -ItemType Directory -Path $dir | Out-Null
        }
        [IO.File]::WriteAllBytes($Dest, $Bytes)
        Good "Wrote $Dest"
        return $true
    }
    catch
    {
        Err "Write failed: $Dest - $( $_.Exception.Message )"; return $false
    }
}

# Rasterize SVG via Svg.Skia + SkiaSharp at given square size with transparent padding
function ConvertFrom-Svg
{
    param([string]$Path, [int]$Target)
    try
    {
        [Reflection.Assembly]::Load("SkiaSharp") | Out-Null
        [Reflection.Assembly]::Load("Svg.Skia")  | Out-Null

        $typeSvg = [Type]::GetType("Svg.Skia.SKSvg, Svg.Skia", $true)
        $svg = [Activator]::CreateInstance($typeSvg)
        [void]$svg.Load($Path)

        $picProp = $typeSvg.GetProperty("Picture")
        $pic = $picProp.GetValue($svg)
        if ($null -eq $pic)
        {
            throw "SVG picture could not be parsed."
        }

        $origW = [int][Math]::Ceiling($pic.CullRect.Width)
        $origH = [int][Math]::Ceiling($pic.CullRect.Height)
        if ($origW -lt 1 -or $origH -lt 1)
        {
            $origW = $origH = $Target
        }

        $scale = [Math]::Min($Target / [double]$origW, $Target / [double]$origH)
        $newW = [int]([double]$origW * $scale)
        $newH = [int]([double]$origH * $scale)
        $offX = [int](($Target - $newW) / 2)
        $offY = [int](($Target - $newH) / 2)

        $bitmap = New-Object SkiaSharp.SKBitmap ($Target, $Target, [SkiaSharp.SKColorType]::Rgba8888, [SkiaSharp.SKAlphaType]::Premul)
        $canvas = New-Object SkiaSharp.SKCanvas ($bitmap)
        $canvas.Clear([SkiaSharp.SKColors]::Transparent)
        $canvas.Translate($offX, $offY)
        $canvas.Scale($scale, $scale)
        $canvas.DrawPicture($pic)
        $canvas.Flush()
        $canvas.Dispose()

        $image = [SkiaSharp.SKImage]::FromBitmap($bitmap)
        $data = $image.Encode([SkiaSharp.SKEncodedImageFormat]::Png, 100)
        $result = $null

        $ms = New-Object System.IO.MemoryStream
        try
        {
            $data.SaveTo($ms)
            $result = $ms.ToArray()
        }
        finally
        {
            $ms.Dispose()
        }

        if ($data -ne $null)
        {
            $data.Dispose()
        }
        if ($image -ne $null)
        {
            $image.Dispose()
        }
        if ($bitmap -ne $null)
        {
            $bitmap.Dispose()
        }
        return $result
    }
    catch
    {
        Err "SVG rasterization failed for $Path - $( $_.Exception.Message )"
        return $null
    }
}

# Rasterize non-SVG images via Magick.NET to a centered square PNG with transparent padding
function ConvertFrom-Raster
{
    param([string]$Path, [int]$Target)
    try
    {
        if (-not ("ImageMagick.MagickImage" -as [type]))
        {
            throw "Magick.NET not available in this session."
        }

        $img = New-Object "ImageMagick.MagickImage" ($Path)
        try
        {
            $ow = [int]$img.Width; $oh = [int]$img.Height
            if ($ow -lt 1 -or $oh -lt 1)
            {
                throw "Invalid image dimensions."
            }

            $scale = [Math]::Min($Target / [double]$ow, $Target / [double]$oh)
            $nw = [int]([double]$ow * $scale)
            $nh = [int]([double]$oh * $scale)

            $img.FilterType = [ImageMagick.FilterType]::Lanczos
            $geometry = New-Object "ImageMagick.MagickGeometry" ($nw, $nh)
            $img.Resize($geometry) | Out-Null

            $canvas = New-Object "ImageMagick.MagickImage" ([ImageMagick.MagickColor]::Transparent, $Target, $Target)
            try
            {
                $canvas.Composite($img, [ImageMagick.Gravity]::Center, 0, 0, [ImageMagick.CompositeOperator]::Over)

                $ms = New-Object System.IO.MemoryStream
                $canvas.Format = "png"
                $canvas.Write($ms)
                return $ms.ToArray()
            }
            finally
            {
                $canvas.Dispose()
            }
        }
        finally
        {
            $img.Dispose()
        }
    }
    catch
    {
        Err "Raster conversion failed for $Path - $( $_.Exception.Message )"
        return $null
    }
}

# WIC fallback using WPF when Magick is not available or fails
function ConvertFrom-WIC
{
    param([string]$Path, [int]$Target)
    try
    {
        Add-Type -AssemblyName PresentationCore, WindowsBase | Out-Null
        $bmp = New-Object System.Windows.Media.Imaging.BitmapImage
        $bmp.BeginInit(); $bmp.CacheOption = "OnLoad"; $bmp.UriSource = [Uri]("file:///$Path");$bmp.EndInit()
        $dv = New-Object System.Windows.Media.DrawingVisual
        $dc = $dv.RenderOpen()
        $brush = New-Object System.Windows.Media.ImageBrush($bmp)
        $brush.Stretch = [System.Windows.Media.Stretch]::Uniform
        $brush.AlignmentX = [System.Windows.Media.AlignmentX]::Center
        $brush.AlignmentY = [System.Windows.Media.AlignmentY]::Center
        $dc.DrawRectangle($brush, $null, (New-Object System.Windows.Rect(0, 0, $Target, $Target)))
        $dc.Close()
        $rtb = New-Object System.Windows.Media.Imaging.RenderTargetBitmap($Target, $Target, 96, 96, [System.Windows.Media.PixelFormats]::Pbgra32)
        $rtb.Render($dv)
        $enc = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
        $enc.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($rtb))
        $ms = New-Object IO.MemoryStream
        $enc.Save($ms)
        $bytes = $ms.ToArray()
        $ms.Dispose()
        return $bytes
    }
    catch
    {
        Err "WIC fallback failed for $Path - $( $_.Exception.Message )"
        return $null
    }
}

# ============================== ICO BUILDER (PNG-compressed entries) ==============================
function New-IcoFromPngMap
{
    <#
    .SYNOPSIS
      Builds a multi-image .ico file from a hashtable of Size->PNG byte[].
    .NOTES
      Writes PNG-compressed icon entries (Vista+). Allows 256px entry (Width/Height byte set to 0).
  #>
    param(
        [Parameter(Mandatory)] [hashtable] $PngBySize,
        [Parameter(Mandatory)] [string]    $DestPath
    )
    # Filter only entries that have non-empty PNG data
    $entries = @()
    foreach ($k in $PngBySize.Keys | Sort-Object -Descending)
    {
        $bytes = $PngBySize[$k]
        if ($bytes -and $bytes.Length -gt 0)
        {
            $entries += [PSCustomObject]@{ Size = [int]$k; Data = $bytes }
        }
    }
    if ($entries.Count -eq 0)
    {
        throw "No icon images to pack."
    }

    $ms = New-Object IO.MemoryStream
    $bw = New-Object IO.BinaryWriter($ms)

    try
    {
        # ICONDIR
        $bw.Write([UInt16]0)   # Reserved
        $bw.Write([UInt16]1)   # Type = 1 (icon)
        $bw.Write([UInt16]$entries.Count) # Count

        # Reserve space for directory; we’ll rewrite offsets after we know data lengths
        $dirStart = $ms.Position
        foreach ($e in $entries)
        {
            # ICONDIRENTRY placeholder (16 bytes each)
            # Width, Height, ColorCount, Reserved, Planes(2), BitCount(2), BytesInRes(4), ImageOffset(4)
            $bw.Write([Byte]([Math]::Min($e.Size, 255)))     # 256 must be 0
            $bw.Write([Byte]([Math]::Min($e.Size, 255)))
            $bw.Write([Byte]0)   # Color count
            $bw.Write([Byte]0)   # Reserved
            $bw.Write([UInt16]0) # Planes (0 for PNG data)
            $bw.Write([UInt16]32)# BitCount (informational)
            $bw.Write([UInt32]0) # BytesInRes placeholder
            $bw.Write([UInt32]0) # Offset placeholder
        }

        # Write image data and collect metadata
        $dataMeta = @()
        foreach ($e in $entries)
        {
            $offset = [UInt32]$ms.Position
            $bw.Write($e.Data)
            $length = [UInt32]$e.Data.Length
            $dataMeta += [PSCustomObject]@{ Size = $e.Size; Offset = $offset; Length = $length }
        }

        # Go back and fill directory with real lengths/offsets
        $ms.Position = $dirStart
        foreach ($meta in $dataMeta)
        {
            $sizeByte = [Byte]([Math]::Min($meta.Size, 255))
            $bw.Write($sizeByte)       # Width
            $bw.Write($sizeByte)       # Height
            $bw.Write([Byte]0)         # ColorCount
            $bw.Write([Byte]0)         # Reserved
            $bw.Write([UInt16]0)       # Planes
            $bw.Write([UInt16]32)      # BitCount
            $bw.Write([UInt32]$meta.Length) # BytesInRes
            $bw.Write([UInt32]$meta.Offset) # ImageOffset
        }

        $bw.Flush()
        $bytesOut = $ms.ToArray()
        $dir = [IO.Path]::GetDirectoryName($DestPath)
        if (-not (Test-Path -LiteralPath $dir))
        {
            New-Item -ItemType Directory -Path $dir | Out-Null
        }
        [IO.File]::WriteAllBytes($DestPath, $bytesOut)
        return $true
    }
    finally
    {
        $bw.Dispose();$ms.Dispose()
    }
}

# ============================== FILE PICKER ==============================
Add-Type -AssemblyName System.Windows.Forms
function Pick-SourceFiles
{
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = "Select one or more images"
    $dlg.Filter = "Images|*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tif;*.tiff;*.tga;*.webp;*.ico;*.svg;*.heic;*.heif|All files|*.*"
    $dlg.Multiselect = $true
    $dlg.CheckFileExists = $true
    $dlg.RestoreDirectory = $true
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
    {
        return $dlg.FileNames
    }
    @()
}

# ============================== MAIN ==============================
function Start-AlkeariIconSuite {
    [CmdletBinding()]
    param()

    $brand = New-AlkeariIconBranding

    Clear-Host
    Write-Host ""
    Write-Host "──────────────────────────────────────────────────────────" -ForegroundColor $Gold
    Write-Host "   Alkeari Labs LLC" -ForegroundColor $Gold
    Write-Host "──────────────────────────────────────────────────────────" -ForegroundColor $Gold
    Write-Host "   Icon Suite Generator (.ico): 256, 128, 64, 32, 16" -ForegroundColor $Silver
    Write-Host "──────────────────────────────────────────────────────────" -ForegroundColor $Gold
    Write-Host ""

    Info "Using dependency cache in: $DepsRoot"
    Info "Bootstrapping local dependencies. First run downloads packages..."
    Initialize-Dependencies

    Info "Pick your source images..."
    $files = Pick-SourceFiles
    if (-not $files -or $files.Count -eq 0)
    {
        Warn "Operation cancelled. No files selected."
        # ============================== END OF SCRIPT ==============================
        Write-Host ""
        Write-Host "──────────────────────────────────────────────" -ForegroundColor Yellow
        Write-Host "  Press Enter to exit..." -ForegroundColor Gray
        Write-Host "──────────────────────────────────────────────" -ForegroundColor Yellow
        Read-Host
        exit
    }

    $ts = Get-Date -Format "yyyyMMdd_HHmmss"
    $totalIcoOk = 0; $totalIcoErr = 0
    $MagickAvailable = ("ImageMagick.MagickImage" -as [type]) -ne $null

    foreach ($src in $files)
    {
        Write-Host ""
        Info "Processing: $([IO.Path]::GetFileName($src) )"
        $outRoot = Join-Path ([IO.Path]::GetDirectoryName($src)) ("IconSuite_{0}" -f $ts)
        $name = [IO.Path]::GetFileNameWithoutExtension($src)
        $icoPath = Join-Path $outRoot ("{0}.ico" -f $name)

        $ext = [IO.Path]::GetExtension($src).ToLowerInvariant()

        # Render all sizes to PNG bytes in memory
        $pngMap = @{ }
        foreach ($sz in $IconSizes)
        {
            $png = $null
            if ($ext -eq ".svg")
            {
                $png = ConvertFrom-Svg -Path $src -Target $sz
            }
            else
            {
                if ($MagickAvailable)
                {
                    $png = ConvertFrom-Raster -Path $src -Target $sz
                }
                if ($null -eq $png)
                {
                    Warn "Using WIC fallback for $([IO.Path]::GetFileName($src) ) at ${sz}px"
                    $png = ConvertFrom-WIC -Path $src -Target $sz
                }
            }
            if ($png -ne $null)
            {
                $pngMap[$sz] = $png
            }
            else
            {
                Err "Failed to render ${sz}px for $([IO.Path]::GetFileName($src) )"
            }
        }

        # Build the .ico from the PNG map
        try
        {
            if (New-IcoFromPngMap -PngBySize $pngMap -DestPath $icoPath)
            {
                Good "Wrote ICO → $icoPath"
                $totalIcoOk++
            }
            else
            {
                Err "Failed to write ICO for $([IO.Path]::GetFileName($src) )"
                $totalIcoErr++
            }
        }
        catch
        {
            Err "ICO build failed for $([IO.Path]::GetFileName($src) ) - $( $_.Exception.Message )"
            $totalIcoErr++
        }
    }

    Write-Host ""
    Write-Host "──────────────── Summary ────────────────" -ForegroundColor $Gold
    Write-Host ("   ICO files written: {0}" -f $totalIcoOk) -ForegroundColor $White
    if ($totalIcoErr -gt 0)
    {
        Write-Host ("   ICO failures:      {0}" -f $totalIcoErr) -ForegroundColor $Red
    }
    else
    {
        Write-Host "   ICO failures:      0" -ForegroundColor $Green
    }
    Write-Host "   Output location(s): IconSuite_<timestamp> folder(s) next to each source image" -ForegroundColor $Cyan
    Warn "If a rare HEIC or other edge format fails on some PCs, re-export to PNG."

    # ============================== END OF SCRIPT ==============================
    Write-Host ""
    Write-Host "──────────────────────────────────────────────" -ForegroundColor Yellow
    Write-Host "  Press Enter to exit..." -ForegroundColor Gray
    Write-Host "──────────────────────────────────────────────" -ForegroundColor Yellow
    Read-Host
}

if ($MyInvocation.InvocationName -ne '.') {
    Start-AlkeariIconSuite
}
