<#
  Alkeari Labs LLC - Folder Inventory & File Tree (Modernized)
  - Shared branding helpers
  - Param-based entry point for optional automation
  - WPF GUI preserved, structure clarified
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Try to import shared theming module if present
$themingModule = Join-Path $PSScriptRoot 'Alkeari.AlkTheming.psm1'
if (Test-Path -LiteralPath $themingModule) {
    Import-Module $themingModule -ErrorAction SilentlyContinue
}

# Fallback theme helper if module not present
if (-not (Get-Command -Name New-AlkeariFolderTheme -ErrorAction SilentlyContinue)) {
    function New-AlkeariFolderTheme {
        [CmdletBinding()]
        param()
        [pscustomobject]@{
            Primary  = '#000000'
            Accent1  = '#DAA520'
            Accent2  = '#8B4513'
            Neutral1 = '#C0C0C0'
            Neutral2 = '#F5F5F5'
        }
    }
}

function Start-AlkeariFolderInventory {
    [CmdletBinding()]
    param(
        [string]$InitialPath,
        [string]$InitialOutputPath
    )

    Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase,System.Drawing,System.Windows.Forms | Out-Null

    Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class WinAPI {
  [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
"@
    $hwnd = [WinAPI]::GetConsoleWindow()
    if ($hwnd -ne [IntPtr]::Zero) { [WinAPI]::ShowWindow($hwnd, 0) | Out-Null }

    $theme = New-AlkeariFolderTheme

    # Branding Colors
    $Color_Primary = $theme.Primary
    $Color_Accent1 = $theme.Accent1
    $Color_Accent2 = $theme.Accent2
    $Color_Neutral1 = $theme.Neutral1
    $Color_Neutral2 = $theme.Neutral2

    # Controls
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Alkeari Labs LLC - Folder Inventory"
        Width="780" Height="520" ResizeMode="NoResize" WindowStartupLocation="CenterScreen"
        Background="$Color_Primary" FontFamily="Segoe UI">
  <Grid Margin="20">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="10"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="10"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="10"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <Border Grid.Row="0" CornerRadius="12" Background="$Color_Accent1" Padding="16">
      <StackPanel>
        <TextBlock Text="Alkeari Labs LLC" FontSize="18" FontWeight="Bold" Foreground="$Color_Primary"/>
        <TextBlock Text="Folder Inventory and File Tree" FontSize="14" Foreground="$Color_Primary"/>
      </StackPanel>
    </Border>

    <StackPanel Grid.Row="2" Orientation="Vertical" Background="$Color_Primary">
      <TextBlock Text="Target folder" Foreground="$Color_Neutral2" Margin="0,0,0,6"/>
      <DockPanel>
        <TextBox Name="txtPath" Foreground="$Color_Neutral2" Background="$Color_Primary" BorderBrush="$Color_Neutral1" Margin="0,0,6,0"/>
        <Button Name="btnBrowsePath" Content="Browse" Background="$Color_Accent2" Foreground="$Color_Neutral2" Padding="8,4"/>
      </DockPanel>

      <TextBlock Text="Output folder" Foreground="$Color_Neutral2" Margin="0,12,0,6"/>
      <DockPanel>
        <TextBox Name="txtOut" Foreground="$Color_Neutral2" Background="$Color_Primary" BorderBrush="$Color_Neutral1" Margin="0,0,6,0"/>
        <Button Name="btnBrowseOut" Content="Browse" Background="$Color_Accent2" Foreground="$Color_Neutral2" Padding="8,4"/>
      </DockPanel>

      <StackPanel Orientation="Horizontal" Margin="0,12,0,0">
        <CheckBox Name="chkHidden" Content="Include Hidden" Foreground="$Color_Neutral2" Margin="0,0,16,0"/>
        <CheckBox Name="chkSystem" Content="Include System" Foreground="$Color_Neutral2" Margin="0,0,16,0"/>
        <CheckBox Name="chkJson" Content="Also export JSON" Foreground="$Color_Neutral2"/>
      </StackPanel>

      <TextBlock Text="YAML is always produced. JSON is optional." Foreground="$Color_Neutral1" Margin="0,6,0,0"/>
    </StackPanel>

    <StackPanel Grid.Row="4" Orientation="Horizontal" VerticalAlignment="Center">
      <ProgressBar Name="pb" Width="520" Height="16" Minimum="0" Maximum="100" Value="0" Foreground="$Color_Accent1"/>
      <Button Name="btnRun" Content="Generate" Background="$Color_Accent1" Foreground="$Color_Primary" Padding="12,6" Margin="12,0,0,0"/>
    </StackPanel>

    <Border Grid.Row="6" BorderBrush="$Color_Neutral1" BorderThickness="1" CornerRadius="10" Padding="10">
      <ScrollViewer VerticalScrollBarVisibility="Auto">
        <TextBlock Name="txtLog" TextWrapping="Wrap" Foreground="$Color_Neutral2"/>
      </ScrollViewer>
    </Border>

    <StackPanel Grid.Row="7" Orientation="Horizontal" HorizontalAlignment="Right">
      <Button Name="btnClose" Content="Close" Background="$Color_Accent2" Foreground="$Color_Neutral2" Padding="10,6"/>
    </StackPanel>
  </Grid>
</Window>
"@

    # Load XAML
    $reader = (New-Object System.Xml.XmlNodeReader $xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)

    # Controls
    $txtPath = $window.FindName("txtPath")
    $btnBrowsePath = $window.FindName("btnBrowsePath")
    $txtOut = $window.FindName("txtOut")
    $btnBrowseOut = $window.FindName("btnBrowseOut")
    $chkHidden = $window.FindName("chkHidden")
    $chkSystem = $window.FindName("chkSystem")
    $chkJson = $window.FindName("chkJson")
    $pb = $window.FindName("pb")
    $btnRun = $window.FindName("btnRun")
    $txtLog = $window.FindName("txtLog")
    $btnClose = $window.FindName("btnClose")

    # Defaults
    $chkJson.IsChecked = $false
    $txtPath.Text = [Environment]::GetFolderPath('MyDocuments')
    $txtOut.Text = $txtPath.Text

    function LogLine([string]$s)
    {
        $txtLog.Text += ("`n" + $s)
    }

    # Pickers
    $btnBrowsePath.Add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = "Select the target folder to scan"
        $dlg.ShowNewFolderButton = $false
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
        {
            $txtPath.Text = $dlg.SelectedPath
        }
    })
    $btnBrowseOut.Add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = "Select the output folder for reports"
        $dlg.ShowNewFolderButton = $true
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
        {
            $txtOut.Text = $dlg.SelectedPath
        }
    })

    # Close
    $btnClose.Add_Click({ $window.Close() })

    # Run
    $btnRun.Add_Click({
        $btnRun.IsEnabled = $false
        $pb.Value = 0
        $txtLog.Text = ""
        try
        {
            $root = $txtPath.Text.Trim()
            $outDir = $txtOut.Text.Trim()
            if ([string]::IsNullOrWhiteSpace($root) -or -not (Test-Path -LiteralPath $root -PathType Container))
            {
                [System.Windows.MessageBox]::Show("Please choose a valid target folder.", "Alkeari Labs LLC")
                $btnRun.IsEnabled = $true
                return
            }
            if ( [string]::IsNullOrWhiteSpace($outDir))
            {
                $outDir = $root
            }
            if (-not (Test-Path -LiteralPath $outDir))
            {
                New-Item -ItemType Directory -Path $outDir -Force | Out-Null
            }

            $ts = Get-Date -Format "yyyyMMdd_HHmmss"
            $yamlPath = Join-Path $outDir ("FolderReport_{0}.yaml" -f $ts)
            $jsonPath = Join-Path $outDir ("FolderReport_{0}.json" -f $ts)
            $treePath = Join-Path $outDir ("FileTree_{0}.txt" -f $ts)

            $includeHidden = [bool]$chkHidden.IsChecked
            $includeSystem = [bool]$chkSystem.IsChecked
            $alsoJson = [bool]$chkJson.IsChecked

            LogLine "Scanning..."
            $pb.Value = 10

            $attrFilter = {
                param($i, $incHidden, $incSystem)
                if (-not $incHidden -and ($i.Attributes -band [IO.FileAttributes]::Hidden))
                {
                    return $false
                }
                if (-not $incSystem -and ($i.Attributes -band [IO.FileAttributes]::System))
                {
                    return $false
                }
                return $true
            }

            $items = Get-ChildItem -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
            $filtered = foreach ($i in $items)
            {
                if (& $attrFilter $i $includeHidden $includeSystem)
                {
                    $i
                }
            }

            $pb.Value = 35
            LogLine "Building records..."

            $records = $filtered | ForEach-Object {
                $isDir = $_.PSIsContainer
                $sizeBytes = if ($isDir)
                {
                    0
                }
                else
                {
                    [int64]($_.Length)
                }

                # Relative path
                $p = $_.FullName
                $rel = if ( $p.StartsWith($root, [System.StringComparison]::OrdinalIgnoreCase))
                {
                    $p.Substring($root.Length).TrimStart('\', '/')
                }
                else
                {
                    $_.Name
                }

                # Safe property reads
                $created = ""
                $modified = ""
                $attrs = ""
                try
                {
                    $created = $_.CreationTimeUtc.ToString("yyyy-MM-dd HH:mm:ss")
                }
                catch
                {
                }
                try
                {
                    $modified = $_.LastWriteTimeUtc.ToString("yyyy-MM-dd HH:mm:ss")
                }
                catch
                {
                }
                try
                {
                    $attrs = $_.Attributes.ToString()
                }
                catch
                {
                }

                [ordered]@{
                    RelativePath = if ($rel)
                    {
                        $rel
                    }
                    else
                    {
                        "."
                    }
                    Name = $_.Name
                    Extension = if ($isDir)
                    {
                        ""
                    }
                    else
                    {
                        $_.Extension
                    }
                    Type = if ($isDir)
                    {
                        "Directory"
                    }
                    else
                    {
                        "File"
                    }
                    SizeBytes = $sizeBytes
                    SizeHuman = if ($isDir)
                    {
                        ""
                    }
                    else
                    {
                        Format-Size $sizeBytes
                    }
                    CreatedUtc = $created
                    ModifiedUtc = $modified
                    Attributes = $attrs
                    FullPath = $_.FullName
                }
            }


            $pb.Value = 55
            LogLine "Generating file tree..."
            $rootName = Split-Path -Leaf -Path $root
            $treeLines = @()
            $treeLines += $rootName
            $treeLines += (New-TreeLines -BasePath $root -AttrFilter { param($x) & $attrFilter $x $includeHidden $includeSystem })
            $treeLines | Out-File -FilePath $treePath -Encoding UTF8

            $pb.Value = 70
            LogLine "Composing YAML..."

            $report = [ordered]@{
                ReportType = "FolderInventory"
                Root = $root
                Generated = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss 'UTC'")
                Summary = [ordered]@{
                    Files = ($records | Where-Object { $_.Type -eq "File" }).Count
                    Directories = ($records | Where-Object { $_.Type -eq "Directory" }).Count
                    TotalBytes = ($records | Measure-Object -Property SizeBytes -Sum).Sum
                    TotalSize = Format-Size ( ($records | Measure-Object -Property SizeBytes -Sum).Sum)
                }
                Items = $records
            }

            (ConvertTo-Yaml $report) | Out-File -FilePath $yamlPath -Encoding UTF8

            $pb.Value = 85
            if ($alsoJson)
            {
                LogLine "Writing JSON..."
                $report | ConvertTo-Json -Depth 20 | Out-File -Path $jsonPath -Encoding UTF8
            }

            $pb.Value = 100
            LogLine "Done."
            LogLine "YAML: $yamlPath"
            if ($alsoJson)
            {
                LogLine "JSON: $jsonPath"
            }
            LogLine "Tree: $treePath"
        }
        catch
        {
            [System.Windows.MessageBox]::Show("Error: $( $_.Exception.Message )", "Alkeari Labs LLC")
        }
        finally
        {
            $btnRun.IsEnabled = $true
        }
    })

    # Show Window
    $window.Topmost = $false
    [void]$window.ShowDialog()
}

if ($MyInvocation.InvocationName -ne '.') {
    Start-AlkeariFolderInventory
}
