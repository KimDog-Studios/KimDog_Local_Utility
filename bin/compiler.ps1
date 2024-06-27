# Check if Winget is installed
$wingetInstalled = $false

# Check if the winget command exists
if ($null -ne (Get-Command -Name winget -ErrorAction SilentlyContinue)) {
    $wingetInstalled = $true
    Write-Output "Winget is already installed."
}

# Install Winget if it's not already installed
if (-not $wingetInstalled) {
    Write-Output "Winget is not installed. Downloading and installing..."

    # Define the URL to the latest stable release of winget-cli
    $wingetUrl = "https://github.com/microsoft/winget-cli/releases/latest/download/Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.appxbundle"

    # Define the path where to save the installer
    $wingetInstallerPath = "$env:TEMP\winget-cli.appxbundle"

    # Download the installer
    Invoke-WebRequest -Uri $wingetUrl -OutFile $wingetInstallerPath

    # Install winget silently
    Start-Process -FilePath "powershell.exe" -ArgumentList "-Command", "Add-AppxPackage -Path '$wingetInstallerPath' -ForceApplicationShutdown" -Wait
    
    # Check if installation was successful
    if ($null -eq (Get-Command -Name winget -ErrorAction SilentlyContinue)) {
        Write-Error "Failed to install Winget."
    }
    else {
        Write-Output "Winget installation completed successfully."
    }

    # Clean up the installer
    Remove-Item -Path $wingetInstallerPath -Force
}

# URL to the configuration file on GitHub
$jsonFileUrl = "https://raw.githubusercontent.com/KimDog-Studios/KimDog_Utility_Main/main/config/config.json"

# Fetch the configuration file directly
try {
    $config = Invoke-RestMethod -Uri $jsonFileUrl
    Write-Host "Config file fetched successfully."
}
catch {
    Write-Host "Failed to fetch config file."
    Write-Host $_.Exception.Message
    exit 1
}

# List of vcredist versions to check
$vcredist_versions = $config.vcredist_versions.Keys

# Function to check if a vcredist is installed
function Is_VcredistInstalled {
    param (
        [string]$version
    )
    $regPath = "HKLM:\SOFTWARE\Classes\Installer\Products\*"
    $keys = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue

    foreach ($key in $keys) {
        if ($key.PSChildName -match $version) {
            return $true
        }
    }
    return $false
}

# Function to check all required vcredists
function Check_AllVcredists {
    param (
        [array]$versions
    )
    $missingVcredists = @()
    foreach ($version in $versions) {
        if (-not (Is_VcredistInstalled -version $version)) {
            $missingVcredists += $version
        }
    }
    return $missingVcredists
}

# List of winget IDs for vcredist versions
$vcredist_winget_ids = $config.vcredist_versions

# Check for missing vcredists
$missingVcredists = Check_AllVcredists -versions $vcredist_versions

if ($missingVcredists.Count -eq 0) {
    Write-Output "All required vcredists are installed."
}
else {
    Write-Output "The following vcredists are missing: $missingVcredists"
    Write-Output "Installing missing vcredists using winget..."

    foreach ($version in $missingVcredists) {
        $wingetId = $vcredist_winget_ids[$version]
        if ($wingetId) {
            Write-Output "Installing vcredist version $version using winget ID $wingetId..."
            Start-Process -FilePath "winget" -ArgumentList "install", "--id", $wingetId, "--silent", "--accept-package-agreements", "--accept-source-agreements" -Wait
        }
        else {
            Write-Output "No winget ID found for vcredist version $version."
        }
    }

    Write-Output "Installation completed. Please verify if all vcredists are installed."
}

Write-Output "Downloading Source Code from GitHub and Grabbing latest Version..."
##Download Source Code

# Define the URL of the file to download
$url = "https://github.com/KimDog-Studios/KimDog_Utility_Main/releases/download/latest/Local.Version.zip"

# Define the base path and folders
$baseFolder = "C:\"
$kimDogStudiosFolder = "KimDog-Studios"
$kimDogUtilityFolder = "KimDog's Utility"
$kimDogStudiosPath = Join-Path -Path $baseFolder -ChildPath $kimDogStudiosFolder
$extractPath = Join-Path -Path $kimDogStudiosPath -ChildPath $kimDogUtilityFolder
$zipFile = Join-Path -Path $kimDogStudiosPath -ChildPath "Local.Version.zip"
$exeFile = Join-Path -Path $extractPath -ChildPath "UtilityGUI.exe"  # Adjust this based on your actual .exe file name

# Function to download the file
function Download-File {
    param (
        [string]$url,
        [string]$output
    )

    try {
        # Create a WebClient object
        $webClient = New-Object System.Net.WebClient

        # Download the file
        $webClient.DownloadFile($url, $output)

        Write-Host "File downloaded successfully to $output"
    }
    catch {
        Write-Error "Failed to download file: $_"
    }
}

# Function to extract the zip file
function Extract-ZipFile {
    param (
        [string]$zipFile,
        [string]$extractPath
    )

    try {
        # Ensure the extraction path exists
        if (-Not (Test-Path -Path $extractPath)) {
            New-Item -ItemType Directory -Path $extractPath | Out-Null
        }

        # Extract the contents of the zip file
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [System.IO.Compression.ZipFile]::ExtractToDirectory($zipFile, $extractPath)

        Write-Host "Files extracted successfully to $extractPath"
    }
    catch {
        Write-Error "Failed to extract files: $_"
    }
}

# Function to create shortcut on Desktop
function Create-DesktopShortcut {
    param (
        [string]$targetPath,
        [string]$shortcutName
    )

    try {
        $desktopPath = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)
        $shortcutFile = Join-Path -Path $desktopPath -ChildPath "$shortcutName.lnk"

        $WScriptShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WScriptShell.CreateShortcut($shortcutFile)
        $Shortcut.TargetPath = $targetPath
        $Shortcut.Save()

        Write-Host "Desktop shortcut created: $shortcutFile"
    }
    catch {
        Write-Error "Failed to create desktop shortcut: $_"
    }
}

# Function to create shortcut in Start Menu
function Create-StartMenuShortcut {
    param (
        [string]$targetPath,
        [string]$shortcutName
    )

    try {
        $programsFolder = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Programs)
        $shortcutFolder = Join-Path -Path $programsFolder -ChildPath "KimDog-Studios"
        $shortcutFile = Join-Path -Path $shortcutFolder -ChildPath "$shortcutName.lnk"

        # Create KimDog-Studios folder if it doesn't exist
        if (-Not (Test-Path -Path $shortcutFolder)) {
            New-Item -ItemType Directory -Path $shortcutFolder | Out-Null
        }

        $WScriptShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WScriptShell.CreateShortcut($shortcutFile)
        $Shortcut.TargetPath = $targetPath
        $Shortcut.Save()

        Write-Host "Start Menu shortcut created: $shortcutFile"
    }
    catch {
        Write-Error "Failed to create Start Menu shortcut: $_"
    }
}

# Function to compare and replace files
function Compare-And-ReplaceFiles {
    param (
        [string]$sourcePath,
        [string]$destinationPath
    )

    try {
        # Get all files in the source directory
        $sourceFiles = Get-ChildItem -Path $sourcePath -Recurse

        foreach ($file in $sourceFiles) {
            $destinationFile = Join-Path -Path $destinationPath -ChildPath $file.FullName.Substring($sourcePath.Length + 1)

            # Compare file versions or hashes to determine if the file needs replacement
            if (!(Test-Path -Path $destinationFile) -or (Get-FileHash -Path $file.FullName).Hash -ne (Get-FileHash -Path $destinationFile).Hash) {
                # Replace the file
                Copy-Item -Path $file.FullName -Destination $destinationFile -Force
                Write-Host "File replaced: $destinationFile"
            }
        }

        Write-Host "Files comparison and replacement completed."
    }
    catch {
        Write-Error "Failed to compare and replace files: $_"
    }
}

# Function to clean up existing KimDog's Utility folder
function Clean-UpKimDogUtility {
    param (
        [string]$path
    )

    try {
        if (Test-Path -Path $path) {
            # Remove all files and subdirectories
            Remove-Item -Path $path\* -Force -Recurse
            Write-Host "Existing files in $path deleted."
        }
    }
    catch {
        Write-Error "Failed to delete existing files: $_"
    }
}

# Clean up existing files in KimDog's Utility folder
Clean-UpKimDogUtility -path $extractPath

# Download the file
Download-File -url $url -output $zipFile

# Extract the contents of the zip file directly to the final destination
Extract-ZipFile -zipFile $zipFile -extractPath $extractPath

# Optionally, remove the zip file after extraction
Remove-Item -Path $zipFile -Force

# Compare and replace files
Compare-And-ReplaceFiles -sourcePath $extractPath -destinationPath $extractPath

# Create shortcuts
if (Test-Path -Path $exeFile) {
    Create-DesktopShortcut -targetPath $exeFile -shortcutName "KimDog's Utility"
    Create-StartMenuShortcut -targetPath $exeFile -shortcutName "KimDog's Utility"
}
else {
    Write-Error "Executable file not found: $exeFile"
}
