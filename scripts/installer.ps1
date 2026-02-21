# Set the current directory to the script's directory
Set-Location -Path $PSScriptRoot

# Set the title of the PowerShell window
$Host.UI.RawUI.WindowTitle = "Windows 21H2to7 Language Pack Installer"

# Variables
$PACKVERSION = "3.6.1"
$LOCALE = (Get-Culture).Name
$rootFolderPath = Split-Path -Parent $PSScriptRoot
$filesFolderPath = Join-Path -Path $rootFolderPath -ChildPath "files"
$registryPath = "HKLM:\SOFTWARE\WOW6432Node\Win10to7"
$registryVersionValueName = "Version"

# Registry key mappings for conditional features
$script:RegistryKeyMappings = @{
    "LegacyUAC"     = "AllowLegacyUACPatch"
    "LegacyCPL"     = "AllowLegacyCPL"
    "VANPatch"      = "AllowVANPatch"
    "UserCPL"       = "AllowUserCPL"
    "MMC"           = "AllowMMCPatch"
    "Accessibility" = "Use7Accessibility"
}

# Branding to SKU folder mappings
$script:BrandingSkuMappings = @{
    "win7pro"       = "Professional"
    "win7home"      = "HomePremium"
    "win7LTSC"      = "PosReady"
    "win7enterprise" = "Enterprise"
    "win7ult"       = "Ultimate"
}

#region Helper Functions

function Get-RegistryKeyForLine {
    <#
    .SYNOPSIS
    Gets the registry key name for a given line based on feature tags.
    #>
    param (
        [string]$line
    )

    foreach ($key in $script:RegistryKeyMappings.Keys) {
        if ($line -match $key) {
            return $script:RegistryKeyMappings[$key]
        }
    }
    return $null
}

function Test-RegistryFeatureEnabled {
    <#
    .SYNOPSIS
    Tests if a registry feature is enabled for the given line.
    .PARAMETER line
    The line from files.txt to check
    .PARAMETER checkValue
    If true, checks the actual registry value. If false, only checks if key exists (for patching phase).
    #>
    param (
        [string]$line,
        [bool]$checkValue = $true
    )

    $registryKey = Get-RegistryKeyForLine -line $line

    if ($registryKey) {
        $registryValue = Get-ItemProperty -Path $registryPath -Name $registryKey -ErrorAction SilentlyContinue

        if ($checkValue) {
            # Copying phase: Check if value is true OR line contains "ALL"
            if (-not $registryValue.$registryKey -and $line -notmatch "ALL") {
                return $false
            }
        } else {
            # Patching phase: Only check if registry key exists
            if ($null -eq $registryValue) {
                return $false
            }
        }
    }
    return $true
}

function Remove-FeatureTagsFromPath {
    <#
    .SYNOPSIS
    Removes feature tags from file paths using the RegistryKeyMappings hashtable.
    #>
    param (
        [string]$filePath
    )

    # Build pattern from RegistryKeyMappings keys plus "ALL"
    $tags = ($script:RegistryKeyMappings.Keys + "ALL") -join '|'
    $pattern = "\s+($tags)"

    return $filePath -replace $pattern, ""
}

function Parse-FileLine {
    <#
    .SYNOPSIS
    Parses a line from files.txt and returns directory and file components.
    #>
    param (
        [string]$line
    )

    if ($line -match '^(?:"([^"]+)"|(\S+))\s+(.*)$') {
        $directoryPart = if ($matches[1] -ne $null) { $matches[1] } else { $matches[2] }
        $file = $matches[3]

        # Replace locale placeholder and convert to lowercase
        $directoryPart = $directoryPart.ToLower() -replace '\$LOCALE', $LOCALE.ToLower()
        $file = $file.ToLower() -replace '\$LOCALE', $LOCALE.ToLower()

        # Clean feature tags from filename using centralized mapping
        $file = Remove-FeatureTagsFromPath -filePath $file

        return @{
            Directory = $directoryPart
            File = $file
            OriginalLine = $line
        }
    }
    return $null
}

function Get-UniqueFileName {
    <#
    .SYNOPSIS
    Generates a unique filename by appending a counter if the file exists.
    #>
    param (
        [string]$filePath
    )

    $counter = 1
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($filePath)
    $extension = [System.IO.Path]::GetExtension($filePath)
    $directory = [System.IO.Path]::GetDirectoryName($filePath)

    while (Test-Path -Path $filePath -PathType Leaf) {
        $filePath = Join-Path -Path $directory -ChildPath "$baseName($counter)$extension"
        $counter++
    }

    return $filePath
}

function Take-OwnershipAndRename {
    <#
    .SYNOPSIS
    Takes ownership of a file and renames it with a backup extension.
    #>
    param (
        [string]$filePath,
        [string]$newName
    )

    $user = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    $takeown = Start-Process "takeown.exe" -ArgumentList "/F `"$filePath`"" -NoNewWindow -Wait -PassThru
    if ($takeown.ExitCode -eq 0) {
        $icacls = Start-Process "icacls.exe" -ArgumentList "`"$filePath`" /grant `"$user`":F" -NoNewWindow -Wait -PassThru
        if ($icacls.ExitCode -eq 0) {
            $newName = Get-UniqueFileName -filePath $newName
            Rename-Item -Path $filePath -NewName $newName -Force
        } else {
            Write-Host "Error: Failed to grant permissions for '$filePath'."
            return $false
        }
    } else {
        Write-Host "Error: Failed to take ownership of '$filePath'."
        return $false
    }

    return $true
}

function Get-BrandingSkuFolder {
    <#
    .SYNOPSIS
    Gets the SKU folder name based on the branding registry value.
    #>

    $brandingKey = Get-ItemProperty -Path $registryPath -Name "Branding" -ErrorAction SilentlyContinue
    if ($brandingKey) {
        return $script:BrandingSkuMappings[$brandingKey.Branding]
    }
    return $null
}

function Get-ResourceFilePath {
    <#
    .SYNOPSIS
    Gets the correct resource file path, handling the special basebrd case.
    #>
    param (
        [string]$directoryPart,
        [string]$fileRes
    )

    # Handle basebrd.dll.mui.res special case
    if ($fileRes -eq "basebrd.dll.mui.res") {
        $skuFolder = Get-BrandingSkuFolder
        if ($skuFolder) {
            $resDirectoryPart = "$directoryPart\$skuFolder"
        } else {
            $resDirectoryPart = $directoryPart
        }
        $filePathWithRes = Join-Path -Path $filesFolderPath -ChildPath "$LOCALE\$resDirectoryPart"
    } else {
        $filePathWithRes = Join-Path -Path $filesFolderPath -ChildPath "$LOCALE\$directoryPart"
    }

    return Join-Path -Path $filePathWithRes -ChildPath $fileRes
}

function Invoke-ResourcePatch {
    <#
    .SYNOPSIS
    Applies a resource patch using ResourceHacker.
    #>
    param (
        [string]$muiFilePath,
        [string]$resourceFilePath
    )

    # Backup and take ownership
    if (-not (Take-OwnershipAndRename -filePath $muiFilePath -newName "$muiFilePath.bak")) {
        Write-Host "Error: Unable to take ownership and rename '$muiFilePath'."
        return $false
    }

    # Get ResourceHacker path
    $resourceHackerPath = Join-Path -Path (Get-Item -Path $PSScriptRoot).Parent.FullName -ChildPath "bin\ResourceHacker.exe"

    # ResourceHacker is x86, needs sysnative for System32
    $backupPath = "$muiFilePath.bak" -replace '\\system32\\', '\\sysnative\\'
    $muiFilePathSysnative = $muiFilePath -replace '\\system32\\', '\\sysnative\\'

    # Build and execute arguments
    $arguments = "-action addoverwrite -open `"$backupPath`" -save `"$muiFilePathSysnative`" -resource `"$resourceFilePath`""
    $process = Start-Process -FilePath $resourceHackerPath -ArgumentList $arguments -NoNewWindow -Wait -PassThru

    return $process.ExitCode -eq 0
}

function Copy-FileWithBackup {
    <#
    .SYNOPSIS
    Copies a file to the destination, backing up the existing file if present.
    #>
    param (
        [string]$sourcePath,
        [string]$destinationPath
    )

    # Ensure source exists
    if (-not (Test-Path -Path $sourcePath -PathType Leaf)) {
        Write-Host "Error: The source file '$sourcePath' does not exist."
        return $false
    }

    # Backup existing file
    if (Test-Path -Path $destinationPath -PathType Leaf) {
        if (-not (Take-OwnershipAndRename -filePath $destinationPath -newName "$destinationPath.bak")) {
            Write-Host "Error: Unable to take ownership and rename '$destinationPath'."
            return $false
        }
    }

    Write-Host "$destinationPath..."
    Copy-Item -Path $sourcePath -Destination $destinationPath -Force
    return $true
}

function Clear-MuiCache {
    <#
    .SYNOPSIS
    Clears the MUI cache to force Windows to use new strings.
    #>
    $MUICache = "HKCU:\SOFTWARE\Classes\Local Settings\MuiCache"
    if (Test-Path -Path $MUICache) {
        Remove-Item -Path $MUICache -Force -Recurse -ErrorAction SilentlyContinue
    }
    New-Item -Path $MUICache -Force -ErrorAction SilentlyContinue | Out-Null
}

#endregion

#region Initialization

# Check if the /Locale argument is specified
foreach ($arg in $args) {
    if ($arg -match "^/Locale=(.+)$") {
        $LOCALE = $matches[1]
    }
}

# Load file list
$FILES = Get-Content -Path "$filesFolderPath\$LOCALE\files.txt"

# Check if the locale folder exists
$localeFolderPath = Join-Path -Path $filesFolderPath -ChildPath $LOCALE
if (-not (Test-Path -Path $localeFolderPath -PathType Container)) {
    Write-Host "Error: The current Locale '$LOCALE' is not supported by the Windows 21H2to7 Language Pack installer. The installer will now close."
    Pause
    Exit
}

# Display header
Clear-Host
Write-Host ""
Write-Host "Language Pack for Windows 21H2to7 Version $PACKVERSION"
Write-Host "=========================================================================="
Write-Host ""

# Check for admin privileges
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

# Self-elevate if needed
if (-not $isAdmin) {
    Write-Host "This script requires administrator privileges. Requesting elevation..."

    # Build the argument list to pass to the new elevated instance
    $argumentList = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""

    # Pass through any existing arguments
    if ($args.Count -gt 0) {
        $argumentList += " " + ($args -join " ")
    }

    # Start new elevated PowerShell process
    Start-Process -FilePath "powershell.exe" -ArgumentList $argumentList -Verb RunAs

    # Exit the non-elevated instance
    Exit
}

# Check version compatibility
$actualRegistryVersion = $null
$registryProperty = Get-ItemProperty -Path $registryPath -Name $registryVersionValueName -ErrorAction SilentlyContinue
if ($registryProperty) {
    $actualRegistryVersion = $registryProperty.$registryVersionValueName
}

if ($actualRegistryVersion -and $PACKVERSION -ne $actualRegistryVersion) {
    Write-Host ""
    Write-Warning "This language pack needs version: $PACKVERSION, but the installed version of the Transformation Pack is: $actualRegistryVersion. The installer can continue, but some features might not work correctly."
    Write-Host ""
    Pause
}

# Display confirmation
Clear-Host
Write-Host ""
Write-Host "Language Pack for Windows 21H2to7 Version $PACKVERSION"
Write-Host "=========================================================================="
Write-Host ""

Write-Host "This script will update the MUIs on non En-US versions of Windows"
Write-Host "for use in the 21H2to7 Transformation Pack."
Write-Host ""
Write-Host "The modifications are community-made and might not be 1:1 to Windows 7."
Write-Host ""
Write-Host "THESE MODIFICATIONS CANNOT BE UNDONE LATER."
Write-Host ""
Write-Host "Confirm that the locale below is the one you want to install:"
Write-Host ""
Write-Host "  Locale to install: $LOCALE"
Write-Host ""
Write-Host "Your Computer will require a restart after the installer finishes copying the files."
Write-Host ""

$choice = Read-Host "Do you wish to continue? Doing so will replace system files. (Y/N)"
if ($choice -ne 'Y') {
    Exit
}

#endregion

#region Installation

Clear-Host

# Remove language pack app
Write-Host "Removing Microsoft.LanguageExperiencePack$LOCALE..."
$packs = Get-AppxPackage -Name "Microsoft.LanguageExperiencePack$LOCALE*"
Write-Host "Found $($packs.Count) language pack(s)"
if ($packs) {
    $packs | Remove-AppxPackage
    Write-Host "Language pack removal completed."
} else {
    Write-Host "No language pack found."
}

# Phase 1: Patch .res files
Write-Host "Patching files..."

foreach ($line in $FILES) {
    $parsedLine = Parse-FileLine -line $line
    if (-not $parsedLine) { continue }

    # Check if feature is enabled (patching phase: only check if registry key exists)
    if (-not (Test-RegistryFeatureEnabled -line $parsedLine.OriginalLine -checkValue $false)) {
        continue
    }

    $file = $parsedLine.File
    $directoryPart = $parsedLine.Directory

    # Skip files that are not .res
    if ($file -notlike "*.res") {
        continue
    }

    $fileRes = $file -replace '\.mui$', '.res'
    $filePathWithRes = Get-ResourceFilePath -directoryPart $directoryPart -fileRes $fileRes

    # Build MUI file path
    $systemDrive = [System.Environment]::GetEnvironmentVariable('SystemDrive')
    $muiFilePath = Join-Path -Path $systemDrive -ChildPath $directoryPart
    $muiFilePath = Join-Path -Path $muiFilePath -ChildPath ($file -replace '\.res$', '')

    # Check for the .res file and patch accordingly
    if (Test-Path -Path $filePathWithRes -PathType Leaf) {
        Invoke-ResourcePatch -muiFilePath $muiFilePath -resourceFilePath $filePathWithRes
    } else {
        Write-Host "Warning: The file '$fileRes' does not exist in the locale folder '$filePathWithRes'."
    }
}

# Phase 2: Copy files
Write-Host "Copying files..."

foreach ($line in $FILES) {
    $parsedLine = Parse-FileLine -line $line
    if (-not $parsedLine) { continue }

    # Check if feature is enabled
    if (-not (Test-RegistryFeatureEnabled -line $parsedLine.OriginalLine)) {
        continue
    }

    $file = $parsedLine.File
    $directoryPart = $parsedLine.Directory

    # Skip .res files
    if ($file -like "*.res") {
        continue
    }

    # Build source and destination paths
    $sourcePath = Join-Path -Path $filesFolderPath -ChildPath "$LOCALE\$directoryPart"
    $sourcePath = Join-Path -Path $sourcePath -ChildPath $file

    $destinationPath = Join-Path -Path $env:SystemDrive -ChildPath $directoryPart
    $destinationPath = Join-Path -Path $destinationPath -ChildPath $file

    # Copy file with backup
    Copy-FileWithBackup -sourcePath $sourcePath -destinationPath $destinationPath
}

Write-Host "Copying files is done."

# Phase 3: Cleanup and restart
Clear-MuiCache

Write-Host ""
Write-Host "The $LOCALE Language Pack for Windows 21H2to7 has been installed! Your Computer will restart to apply the changes."
Write-Host "Original files have been backed up with .bak extension."
Write-Host ""
Pause

for ($i = 10; $i -ge 1; $i--) {
    Write-Host "Your Computer will restart in $i seconds..."
    Start-Sleep -Seconds 1
}
Restart-Computer

#endregion
