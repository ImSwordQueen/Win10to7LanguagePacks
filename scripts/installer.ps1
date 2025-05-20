# Set the current directory to the script's directory
Set-Location -Path $PSScriptRoot

# Set the title of the PowerShell window
$Host.UI.RawUI.WindowTitle = "Windows 21H2to7 Language Pack Installer"

# Variables
$PACKVERSION = "3.6.1"
$LOCALE = (Get-Culture).Name
$rootFolderPath = Split-Path -Parent $PSScriptRoot
$filesFolderPath = Join-Path -Path $rootFolderPath -ChildPath "files"
$FILES = Get-Content -Path "$filesFolderPath\$LOCALE\files.txt"
$registryPath = "HKLM:\SOFTWARE\WOW6432Node\Win10to7"
$registryVersionValueName = "Version"

# Check if the /Locale argument is specified
foreach ($arg in $args) {
    if ($arg -match "^/Locale=(.+)$") {
        $LOCALE = $matches[1]
    }
}

# Check if the locale folder exists in the windows directory
$windowsFolderPath = Join-Path -Path $rootFolderPath -ChildPath "files"
$localeFolderPath = Join-Path -Path $windowsFolderPath -ChildPath $LOCALE

if (-not (Test-Path -Path $localeFolderPath -PathType Container)) {
    Write-Host "Error: The current Locale '$LOCALE' is not supported by the Windows 21H2to7 Language Pack installer. The installer will now close."
    Pause
    Exit
}

# Texts
Clear-Host
Write-Host ""
Write-Host "Language Pack for Windows 21H2to7 Version $PACKVERSION"
Write-Host "=========================================================================="
Write-Host ""

# Check for admin (If for some reason the person runs installer.ps1 instead of Install.CMD)
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "You must run this script as administrator. Right click the script and click"
    Write-Host "Run as Administrator, and if you need to pass arguments, run it in a Command"
    Write-Host "Prompt window that is running as administrator."
    Write-Host ""
    Pause
    Exit
}

# Check if the version of 21H2to7 is the same as the language pack installer.
$actualRegistryVersion = (Get-ItemProperty -Path $registryPath -Name $registryVersionValueName).$registryVersionValueName
if ($PACKVERSION -ne $actualRegistryVersion) {
	Write-Host ""
    Write-Warning "This language pack needs version: $PACKVERSION, but the installed version of the Transformation Pack is: $actualRegistryVersion. The installer can continue, but some features might not work correctly."
	Write-Host ""
    Pause
}

# Continue if everything is good.
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

# INSTALL
Clear-Host

# Remove The language pack "app" for that specific locale.
Write-Host "Removing Microsoft.LanguageExperiencePack$locale..."
Get-AppxPackage -Name "Microsoft.LanguageExperiencePack$locale*" | Remove-AppxPackage

# Files
Write-Host "Patching files..."

# Remove the checks from the filenames.
function Clean-FilePath {
    param (
        [string]$filePath
    )
    return $filePath -replace "\s+ALL|\s+LegacyCPL|\s+LegacyUAC|\s+VANPatch|\s+UserCPL|\s+MMC|\s+Accessibility", ""
}

function Take-OwnershipAndRename {
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

function Get-UniqueFileName {
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

function Rename-FileWithBackup {
    param (
        [string]$filePath
    )

    $backupPath = Get-UniqueFileName -filePath "$filePath.bak"
    Rename-Item -Path $filePath -NewName $backupPath -Force
    return $backupPath
}

foreach ($line in $FILES) {
    if ($line -match '^(?:"([^"]+)"|(\S+))\s+(.*)$') {
        if ($matches[1] -ne $null) {
            $directoryPart = $matches[1]
        } else {
            $directoryPart = $matches[2]
        }
        $file = $matches[3]
    } else {
        continue
    }

    $directoryPart = $directoryPart.ToLower() -replace '\$LOCALE', $LOCALE.ToLower()
    $file = $file.ToLower() -replace '\$LOCALE', $LOCALE.ToLower()
    $file = Clean-FilePath -filePath $file
    $skipLine = $false

	switch -Regex ($line) {
        "LegacyUAC" {
            $registryKey = "AllowLegacyUACPatch"
        }
        "LegacyCPL" {
            $registryKey = "AllowLegacyCPL"
        }
        "VANPatch" {
            $registryKey = "AllowVANPatch"
        }
        "UserCPL" {
            $registryKey = "AllowUserCPL"
        }
        "MMC" {
            $registryKey = "AllowMMCPatch"
        }
        "Accessibility" {
            $registryKey = "Use7Accessibility"
        }
        default {
            $registryKey = $null
        }
    }

    if ($registryKey) {
        $registryValue = Get-ItemProperty -Path $registryPath -Name $registryKey -ErrorAction SilentlyContinue

        if ($null -eq $registryValue) {
            $skipLine = $true
        }
    }

    if ($skipLine) {
        continue
    }

    # Skip files that are not .res (We are patching files, not moving)
    if ($file -notlike "*.res") {
        continue
    }

    $fileRes = $file -replace '\.mui$', '.res'
    $filePath = Join-Path -Path $filesFolderPath -ChildPath "$LOCALE\$directoryPart"

    # Check what SKU the user has chosen in the main 21H2to7 installer and then patch it accordingly. (Win10 doesn't get patched (Unless someone adds the basebrd.dll.mui.res in the main directory))
    if ($file -eq "basebrd.dll.mui.res") {
        $brandingKey = Get-ItemProperty -Path "$registryPath" -Name "Branding" -ErrorAction SilentlyContinue
        if ($brandingKey) {
            $muifilepath = Join-Path -Path $directoryPart -ChildPath "basebrd.dll.mui"
            
            # Process the SKU Folder per Branding values.
            if ($file -eq "basebrd.dll.mui.res") {
                $resDirectoryPart = $directoryPart
				switch ($brandingKey.Branding) {
                    "win7pro" { 
                        $resDirectoryPart = Join-Path -Path $resDirectoryPart -ChildPath "Professional"
                    }
                    "win7home" { 
                        $resDirectoryPart = Join-Path -Path $resDirectoryPart -ChildPath "HomePremium"
                    }
                    "win7LTSC" { 
                        $resDirectoryPart = Join-Path -Path $resDirectoryPart -ChildPath "PosReady"
                    }
                    "win7enterprise" { 
                        $resDirectoryPart = Join-Path -Path $resDirectoryPart -ChildPath "Enterprise"
                    }
                    "win7ult" { 
                        $resDirectoryPart = Join-Path -Path $resDirectoryPart -ChildPath "Ultimate"
                    }
                    default { 
                    }
                }
                $filePathWithRes = Join-Path -Path $filesFolderPath -ChildPath "$LOCALE\$resDirectoryPart"
                $filePathWithRes = Join-Path -Path $filePathWithRes -ChildPath "basebrd.dll.mui.res"
            }
        }
    }

    # Basebrd is a weird case. It's the only file that dynamically changes depending on the registry so it requires more work to get it working like the other files.
    if ($file -eq "basebrd.dll.mui.res") {
        $filePathWithRes = Join-Path -Path $filesFolderPath -ChildPath "$LOCALE\$resDirectoryPart"
        $filePathWithRes = Join-Path -Path $filePathWithRes -ChildPath $fileRes
    } else {
        $filePathWithRes = Join-Path -Path $filesFolderPath -ChildPath "$LOCALE\$directoryPart"
        $filePathWithRes = Join-Path -Path $filePathWithRes -ChildPath $fileRes
    }
    if (-not $filePathWithRes.StartsWith($env:SystemDrive)) {
        $filePathWithRes = Join-Path -Path $env:SystemDrive -ChildPath $filePathWithRes
    }
    $systemDrive = [System.Environment]::GetEnvironmentVariable('SystemDrive')
    $muiFilePath = Join-Path -Path $systemDrive -ChildPath $directoryPart
    $muiFilePath = Join-Path -Path $muiFilePath -ChildPath ($file -replace '\.res$', '')

    # Check for the .res file and patches the file it needs to patch accordingly 
    if (Test-Path -Path $filePathWithRes -PathType Leaf) {
        # Make sure that permissions aren't a problem.
        if (-not (Take-OwnershipAndRename -filePath $muiFilePath -newName "$muiFilePath.bak")) {
            Write-Host "Error: Unable to take ownership and rename '$muiFilePath'."
            continue
        }

        # The ResourceHacker Stage
        $resourceHackerPath = Join-Path -Path (Get-Item -Path $PSScriptRoot).Parent.FullName -ChildPath "bin\ResourceHacker.exe"

        # ResourceHacker is a x86 application so it needs conversion technology to patch System32
        $backupPath = "$muiFilePath.bak" -replace '\\system32\\', '\\sysnative\\'
        $muiFilePath = $muiFilePath -replace '\\system32\\', '\\sysnative\\'

        # Build the arguments for Resource Hacker
        $arguments = "-action addoverwrite -open `"$backupPath`" -save `"$muiFilePath`" -resource `"$filePathWithRes`""

        $process = Start-Process -FilePath $resourceHackerPath -ArgumentList $arguments -NoNewWindow -Wait -PassThru
    } else {
	        Write-Host "Warning: The file '$fileRes' does not exist in the locale folder '$filePathWithRes'."
    }
}

# Files
Write-Host "Copying files..."

# Copy the new files
foreach ($line in $FILES) {
    if ($line -match '^(?:"([^"]+)"|(\S+))\s+(.*)$') {
        if ($matches[1] -ne $null) {
            $directoryPart = $matches[1]
        } else {
            $directoryPart = $matches[2]
        }
        $file = $matches[3]
    }

    $directoryPart = $directoryPart.ToLower() -replace '\$LOCALE', $LOCALE.ToLower()
    $file = $file.ToLower() -replace '\$LOCALE', $LOCALE.ToLower()

    # Clean the file path
    $file = Clean-FilePath -filePath $file

    $filePath = Join-Path -Path $filesFolderPath -ChildPath "$LOCALE\$directoryPart"
    $filePath = Join-Path -Path $filePath -ChildPath $file

    # Ensure the file is there (It happens to the best of us)
    if (-not (Test-Path -Path $filePath -PathType Leaf)) {
        Write-Host "Error: The source file '$filePath' does not exist."
        continue
    }
    
    $skipLine = $false
    switch -Regex ($line) {
        "LegacyUAC" {
            $registryKey = "AllowLegacyUACPatch"
        }
        "LegacyCPL" {
            $registryKey = "AllowLegacyCPL"
        }
        "VANPatch" {
            $registryKey = "AllowVANPatch"
        }
        "UserCPL" {
            $registryKey = "AllowUserCPL"
        }
        "MMC" {
            $registryKey = "AllowMMCPatch"
        }
        "Accessibility" {
            $registryKey = "Use7Accessibility"
        }
        default {
            $registryKey = $null
        }
    }

    if ($registryKey) {
        $registryValue = Get-ItemProperty -Path "$registryPath" -Name $registryKey -ErrorAction SilentlyContinue
        if (-not $registryValue.$registryKey -and $line -notmatch "ALL") {
            $skipLine = $true
        }
    }

    if ($skipLine) {
        continue
    }

    # Skip .res Files (We are moving files, not patching)
    if ($file -like "*.res") {
        continue
    }

    $root = "$env:SystemDrive\"
    $destinationPath = Join-Path -Path $root -ChildPath $directoryPart
    $destinationPath = Join-Path -Path $destinationPath -ChildPath $file

    # Check if the file exists in the destination folder before renaming
    if (Test-Path -Path $destinationPath -PathType Leaf) {
        if (-not (Take-OwnershipAndRename -filePath $destinationPath -newName "$destinationPath.bak")) {
            Write-Host "Error: Unable to take ownership and rename '$destinationPath'."
            continue
        }
    }

    Write-Host "$destinationPath..."
    Copy-Item -Path $filePath -Destination $destinationPath -Force
}
Write-Host "Copying files is done."
	
# Cleanup the MUICache so Windows doesn't reuse pre-existing strings and instead it uses the new ones.

	$MUICache = "HKCU\SOFTWARE\Classes\Local Settings\MuiCache"
	Remove-Item -Path "Registry::$MUICache" -Force -Recurse
	New-Item -Path "Registry::$MUICache" -Force

# Setup is done but not all changes will be applied without a restart so we gonna restart the PC.
Write-Host ""
Write-Host "The $LOCALE Language Pack for Windows 21H2to7 has been installed! Your Computer will restart to apply the changes."
Write-Host ""
Pause

for ($i = 10; $i -ge 1; $i--) {
    Write-Host "Your Computer will restart in $i seconds..."
    Start-Sleep -Seconds 1
}
Restart-Computer

