<#
.SYNOPSIS
    This script installs all Fonts located in the specified destination path or Url. (Needs Admin rights)
.DESCRIPTION
    Script to install all fonts with .ttf or .otf extensions found in the specified directory.
    URL can be used to download a font file.

    Note: 
    By default, if no switch is provided, the script will install all fonts in the destination folder.
    User must restart the application to see the changes or logoff/logon.
    
.NOTES
    Script name: Install-Font-RunAs.ps1
    Author: ehrynkiw@it-ed.com
    Date: 04/25/2024
    Version: 1.1.3
.EXAMPLE
    Usage: .\Install-Font-RunAs.ps1 [-Url <url>] [-Destination <path>] [-Help]

    Optional parameters:
    
    -Destination <path> : The path to the font files to install. (defaults to C:\ited\Fonts)
    -Url <url> : The URL to download the font from.    

# Requirements - Administrator Rights
#>

param (
    [string]$Url,
    [string]$Destination = "C:\ited\Fonts",
    [switch]$Help
)

$fontFolder = 'C:\Windows\Fonts'

function Show-Help {
    Write-Host ""
    Write-Host "Usage: .\Install-Font-RunAs.ps1 [-FontUrl <url>] [-FontName <name>] [-Install] [-Help]"
    Write-Host ""
    Write-Host "Parameters:"
    Write-Host ""
    Write-Host "  -Url <url>          : The URL of the font to install."
    Write-Host "  -Destination <path> : The path to the font files to install. (defaults to C:\ited\Fonts)"
    Write-Host "  -Help               : Show this help message."
    Write-Host ""
    Write-Host "Examples:"
    Write-Host ""
    Write-Host "  .\Install-Font-RunAs.ps1 -FontUrl 'https://example.com/font.ttf' -Install"
    Write-Host "     Installs the font from the specified URL."
    Write-Host ""
}

# Start transcript
Start-Transcript -Path "$Destination\Log\Font-Install.log"

if (-not $Url) {
    # Install all fonts in the destination folder with .ttf or .otf extensions
    $fonts = Get-ChildItem -Path $Destination -Include '*.ttf', '*.otf'
    foreach ($font in $fonts) {
        $targetFontPath = Join-Path $fontFolder $font.Name

        # If the font does not exist in the target folder, copy the font
        if (!(Test-Path $targetFontPath)) {
            Copy-Item $font.FullName -Destination $fontFolder -Force

            # Change the registry based on the font type
            switch ($font.extension) {
                '.otf' { 
                    New-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts' -Name "$($font.Name -replace ".{4}$") (OpenType)" -Type String -Value $font.Name -Force
                }
                '.ttf' {
                    New-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts' -Name "$($font.Name -replace ".{4}$") (TrueType)" -Type String -Value $font.Name -Force
                }
            }
        }
    }

    Write-Host ""
    Write-Host "All Fonts installed successfully."
    Write-Host ""
}
else {
    # Use Invoke-WebRequest to download the file
    try {
        $FontFileName = Split-Path $Url -Leaf
        $FontDestination = Join-Path $Destination $FontFileName
        Invoke-WebRequest -Uri $Url -OutFile $FontDestination

        Write-Host ""
        Write-Host "$FontFileName downloaded to: $FontDestination"

        # Process the downloaded font
        $fonts = Get-ChildItem -Path $Destination -Filter $FontFileName
        foreach ($font in $fonts) {
            $targetFontPath = Join-Path $fontFolder $font.Name

            # If the font does not exist in the target folder, copy the font
            if (!(Test-Path $targetFontPath)) {
                Copy-Item $font.FullName -Destination $fontFolder -Force

                # Change the registry based on the font type
                switch ($font.extension) {
                    '.otf' { 
                        New-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts' -Name "$($font.Name -replace ".{4}$") (OpenType)" -Type String -Value $font.Name -Force
                    }
                    '.ttf' {
                        New-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts' -Name "$($font.Name -replace ".{4}$") (TrueType)" -Type String -Value $font.Name -Force
                    }
                }
            }
        }

        Write-Host ""
        Write-Host "$Font Font installed successfully."
        Write-Host ""

    }
    catch {
        Write-Error "An error occurred while downloading the file: $($_.Exception.Message)"
    }
}

# Stop transcript
Stop-Transcript
