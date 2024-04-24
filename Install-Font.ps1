<#
.SYNOPSIS
	This script installs a custom Windows Font. (Does not need admin rights)
.DESCRIPTION
    	Script to install a custom Windows 10/11 Font by file location or url.
.NOTES
	Script name: Install-Font.ps1
	Author: ehrynkiw@it-ed.com
	Date: 04/23/2024
	Version:  1.0.3  
.EXAMPLE
    	Usage: Usage: .\Install-Font.ps1 [-FontUrl <url>] [-FontName <name>] [-Install] [-Help]
#>

param (
    [string]$FontUrl = "",
    [string]$FontName = "",
    [switch]$Install,
    [switch]$Help
)

function Invoke-FontDownload {
    param (
        [string]$Url,
        [string]$Destination
    )

    $webClient = New-Object System.Net.WebClient
    $webClient.DownloadFile($Url, $Destination)
}

function Show-Help {
    Write-Host ""
    Write-Host "Usage: .\Install-Font.ps1 [-FontUrl <url>] [-FontName <name>] [-Install] [-Help]"
    Write-Host ""
    Write-Host "Parameters:"
    Write-Host ""
    Write-Host "  -FontUrl <url>    : The URL of the font to install."
    Write-Host "  -FontName <name>  : The name of the font to uninstall."
    Write-Host "  -Install          : Install the font."
    Write-Host "  -Help             : Show this help message."
    Write-Host ""
    Write-Host "Examples:"
    Write-Host ""
    Write-Host "  .\Install-Font.ps1 -FontUrl 'https://example.com/font.ttf' -Install"
    Write-Host "     Installs the font from the specified URL."
    Write-Host ""
}

$TempFolder = "C:\Windows\Temp\Fonts"
$SystemFontsPath = "C:\Windows\Fonts"
$UserFontsPath = "$env:LOCALAPPDATA\Microsoft\Windows\Fonts"

if ($Help) {
    Show-Help
    exit
}

if ($Install) {
    if (-not $FontUrl) {
        Write-Error "Font URL is missing. Please provide a valid URL using -FontUrl parameter."
        exit 1
    }

    $FontFileName = Split-Path $FontUrl -Leaf
    $FontDestination = Join-Path $TempFolder $FontFileName

    # Check if font already exists in system-wide fonts directory
    if (Test-Path "$SystemFontsPath\$FontFileName") {
        Write-Host "Font '$FontFileName' already exists in system-wide fonts directory. Skipping installation."
        exit
    }

    # Check if font already exists in user-specific fonts directory
    if (Test-Path "$UserFontsPath\$FontFileName") {
        Write-Host "Font '$FontFileName' already exists in user-specific fonts directory. Skipping installation."
        exit
    }

    # Download font from URL
    Invoke-FontDownload -Url $FontUrl -Destination $FontDestination

    # Install font
    $ShellApp = New-Object -ComObject Shell.Application
    $Destination = $ShellApp.Namespace(0x14)
    $Destination.CopyHere($FontDestination, 0x10)

    Write-Host "Font installed successfully: $FontFileName"

    # Delete temporary copy of font
    Remove-Item $FontDestination -Force

    exit
}

Show-Help
