# Alkeari Labs LLC - Shared Theming and Console Helpers

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

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

Export-ModuleMember -Function New-AlkeariIconBranding,New-AlkeariFolderTheme,New-AlkeariModInstallerTheme,New-AlkeariConsoleTheme

