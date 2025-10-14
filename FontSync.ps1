param (
    [string]$args
)

# --- Load configuration ---
$configPath = "C:\FontSync\config.json"
if (-not (Test-Path $configPath)) {
    Write-Error "Config file not found: $configPath"
    exit 1
}

$config = Get-Content $configPath | ConvertFrom-Json


$fontDest = $config.WindowsFontFolder
$logFile = $config.LogFile
$relativePath = $config.DropboxFontFolder

Add-Type -AssemblyName System.Drawing

Start-Transcript -Path $logFile -Append

function Log {
    param (
        [string]$message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - $message"
    Write-Output $logMessage
}

# --- Check for administrative privileges ---
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Log "This script requires administrative privileges. Please run as administrator."
    Stop-Transcript
    Start-Process powershell.exe "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# --- Function: to check if a font is already installed ---
function IsFont-Installed {
    param (
        [string]$fontName,
        [string]$fontType
    )
    $fonts = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
    foreach ($font in $fonts.PSObject.Properties) {
        if ($font.Name -match [regex]::Escape($fontName) -and $font.Value -like "*.$fontType") {
            return $true
        }
    }
    return $false
}

# --- Function: to get the original font naming ---
function Get-FontDisplayName {
    param([string]$FontPath)

    $ext = [System.IO.Path]::GetExtension($FontPath).ToLowerInvariant()

    try {
        switch ($ext) {
            ".ttf" { 
                $collection = New-Object System.Drawing.Text.PrivateFontCollection
                $collection.AddFontFile($FontPath)
                return $collection.Families[0].Name
            }
            ".otf" {
                $collection = New-Object System.Drawing.Text.PrivateFontCollection
                $collection.AddFontFile($FontPath)
                return $collection.Families[0].Name
            }
            ".ttc" {
                # TTC fonts may contain multiple faces; we’ll list all family names
                $collection = New-Object System.Drawing.Text.PrivateFontCollection
                $collection.AddFontFile($FontPath)
                return ($collection.Families | ForEach-Object { $_.Name }) -join ", "
            }
            ".fon" {
                # FON fonts: no Family info via System.Drawing, use filename as fallback
                return [System.IO.Path]::GetFileNameWithoutExtension($FontPath)
            }
            default {
                return $null
            }
        }
    }
    catch {
        Log "Failed to extract name for '$FontPath': $_"
        return [System.IO.Path]::GetFileNameWithoutExtension($FontPath)
    }
}

# --- Function: to install a font ---
function Install-Font {
    param([System.IO.FileInfo]$fontFile)

    $fontName = Get-FontDisplayName $fontFile.FullName
    $fontType = $fontFile.Extension.ToLowerInvariant().TrimStart('.')

    $typeLabel = switch ($fontType) {
                "otf" { "(OpenType)" }
                "ttf" { "(TrueType)" }
                "ttc" { "(TrueType Collection)" }
                "fon" { "(All Res)" }
                default { "(Font)" }
            }

    # --- Check if the font is already installed with the same type ---
    if (-not (IsFont-Installed -fontName $fontName -fontType $fontType)) {
        try {
            # --- Copy the font file to the Fonts directory ---
            $targetFontPath = Join-Path $fontDest $fontFile.Name
            if (!(Test-Path $targetFontPath)) { 
                Copy-Item -Path $fontFile -Destination $fontDest -Force
            }

            $fontRegistryKey = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
            $fontValue = [System.IO.Path]::GetFileName($fontFile)

            # --- Handle potential special cases for registry value format   ---
            $registryName = "$fontName $typeLabel"

            New-ItemProperty -Path $fontRegistryKey -Name $registryName -Value $fontValue -PropertyType String -Force | Out-Null
            Log "Installed font: $fontName (Type: $fontType)"

            # --- Verify installation ---
            if (Verify-FontInstallation $registryName) {
                Log "Verified installation of font: $fontName (Type: $fontType)"
            } else {
                Log "Warning: Font installation verification failed for $fontName (Type: $fontType)"
            }
        } catch {
            Log "Failed to install font: $fontName (Type: $fontType). Error: $_"
        }
    } else {
        Log "Font already installed: $fontName (Type: $fontType)"
    }
}

# --- Function to verify font installation ---
function Verify-FontInstallation {
    param (
        [string]$registryName
    )

    $installedFont = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts" -Name $registryName -ErrorAction SilentlyContinue
    return $null -ne $installedFont
}

# --- Function to get all unique font names with priority to .otf files ---
function Get-UniqueFonts {
    param (
        [string]$fontsPath
    )

    $fontFiles = Get-ChildItem -Path $fontsPath -Recurse -Include *.ttf, *.otf, *.fon, *.ttc -File
    $fontDict = @{}

    foreach ($fontFile in $fontFiles) {
        $fontName = $fontFile.BaseName
        $fontType = $fontFile.Extension.TrimStart('.')
        $key = "$fontName-$fontType"
       
        if (-not $fontDict.ContainsKey($key)) {
            $fontDict[$key] = $fontFile.FullName
        } elseif ($fontType -eq "otf" -and $fontDict[$key].EndsWith(".ttf")) {
            # --- Prioritize .otf over .ttf ---
            $fontDict[$key] = $fontFile.FullName
        }
    }
    return $fontDict.Values
}

# --- Function to uninstall fonts ---
function Uninstall-Fonts {
    param (
        [string]$fontsPath
    )
    $fontFiles = Get-ChildItem -Path $fontsPath -Recurse -Include *.ttf, *.otf, *.fon, *.ttc -File

    foreach ($fontFile in $fontFiles) {
        $fontName = Get-FontDisplayName $fontFile.FullName
        $fontType = $fontFile.Extension.ToLowerInvariant().TrimStart('.')
        $fontRegistryKey = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"

        $typeLabel = switch ($fontType) {
                "otf" { "(OpenType)" }
                "ttf" { "(TrueType)" }
                "ttc" { "(TrueType Collection)" }
                "fon" { "(All Res)" }
                default { "(Font)" }
            }

        # --- Determine the correct registry name ---
        $registryName = "$fontName $typeLabel"        

        # --- Remove font registry entry ---
        Remove-ItemProperty -Path $fontRegistryKey -Name $registryName -ErrorAction SilentlyContinue


        # --- Remove font file from Fonts directory ---
        $installedFontFile = "$fontDest\$($fontFile.Name)"
        if (Test-Path -Path $installedFontFile) {
            Remove-Item -Path $installedFontFile -Force -ErrorAction SilentlyContinue
            Log "Uninstalled font: $fontName (Type: $fontType)"
        } else {
            Log "Font file not found for uninstallation: $fontName (Type: $fontType)"
        }
    }
}

# --- Function to display progress ---
function Show-Progress {
    param (
        [int]$current,
        [int]$total
    )
    $percentComplete = [math]::Round(($current / $total) * 100)
    Write-Progress -Activity "Installing Fonts" -Status "$percentComplete% Complete" -PercentComplete $percentComplete
}

# Function: get file from a cloudStorage
function Get-CloudStorage {
    param(
        [string]$relativePath
    )
    
    $drives = Get-PSDrive -PSProvider FileSystem | Select-Object -ExpandProperty Root

    foreach ($drive in $drives) {
        $testPath = Join-Path -Path $drive -ChildPath $relativePath
        if (Test-Path $testPath) {
            return $testPath
        }
    }
    return $false
}


# --- Main execution block --
try {
    Log "Starting font installation process..."

    $fontsPath = Get-CloudStorage -relativePath $relativePath

    # --- Check if fonts directory exists ---
    if (-not $fontsPath) {
        throw "Fonts directory not found: $fontsPath"
    }

    # --- Get all unique font files with priority to .otf ---
    $fontFiles = Get-UniqueFonts -fontsPath $fontsPath

    if ($fontFiles.Count -eq 0) {
        Log "No font files found in the specified directory."
    } else {
        $totalFonts = $fontFiles.Count
        $currentFont = 0

        foreach ($fontFile in $fontFiles) {
            $currentFont++
            Show-Progress -current $currentFont -total $totalFonts
            Install-Font -fontFile $fontFile
        }
    }

    Log "Font installation process completed."

    # --- Uncomment the following lines to enable font uninstallation --
    if ($args -ne '-s') {
        $uninstall = Read-Host "Do you want to uninstall the fonts? (Y/N)"
        if ($uninstall -eq "Y") {
            Log "Starting font uninstallation process..."
            Uninstall-Fonts -fontsPath $fontsPath
            Log "Font uninstallation process completed."
        }
    }
    

} catch {
    Log "An error occurred during script execution: $_"
} finally {
    # Clear the progress bar
    Write-Progress -Activity "Installing Fonts" -Completed
    Stop-Transcript
    exit
}