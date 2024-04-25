<#
.SYNOPSIS
    Ce script installe toutes les polices situées dans le chemin de destination spécifié ou à partir d'une URL. (Nécessite des droits d'administrateur)
.DESCRIPTION
    Script pour installer toutes les polices avec les extensions .ttf ou .otf trouvées dans le répertoire spécifié.
    Une URL peut être utilisée pour télécharger un fichier de police.

    Remarque :
    Par défaut, si aucun switch n'est fourni, le script installera toutes les polices dans le dossier de destination.
    L'utilisateur doit redémarrer l'application pour voir les modifications ou se déconnecter/reconnecter.
.NOTES
    Nom du script : Install-Font-RunAs.ps1
    Auteur : ehrynkiw@it-ed.com
    Date : 25/04/2024
    Version : 1.1.3
.EXAMPLE
    Utilisation : .\Install-Font-RunAs.ps1 [-Url <url>] [-Destination <path>] [-Help]

    Paramètres facultatifs :
    
    -Destination <path> : Le chemin vers les fichiers de police à installer. (par défaut : C:\ited\Fonts)
    -Url <url> : L'URL pour télécharger la police.

# Requirements - Droits d'administrateur
#>

param (
    [string]$Url,
    [string]$Destination = "C:\ited\Fonts",
    [switch]$Help
)

$fontFolder = 'C:\Windows\Fonts'

function Show-Help {
    Write-Host ""
    Write-Host "Usage: .\Install-Font-RunAs.ps1 [-Url <url>] [-Destination <path>] [-Help]"
    Write-Host ""
    Write-Host "Parameters:"
    Write-Host ""
    Write-Host "  -Url <url>          : The URL of the font to install."
    Write-Host "  -Destination <path> : The path to the font files to install. (defaults to C:\ited\Fonts)"
    Write-Host "  -Help               : Show this help message."
    Write-Host ""
    Write-Host "Examples:"
    Write-Host ""
    Write-Host "  .\Install-Font-RunAs.ps1 -Url 'https://example.com/font.ttf' "
    Write-Host "     Installs the font from the specified URL."
    Write-Host ""
}

# Start transcript
Start-Transcript -Path "$Destination\Log\Font-Install.log"

if ($Help) {
    Show-Help
}
elseif (-not $Url) {
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
