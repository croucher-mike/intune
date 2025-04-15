<#
.SYNOPSIS
    Converts MSI files to Intunewin format using IntuneWinAppUtil.exe.

.DESCRIPTION
    This script processes MSI files in a specified folder and converts them to Intunewin format
    using the Microsoft Win32 Content Prep Tool (IntuneWinAppUtil.exe).

.PARAMETERS
    None

.EXAMPLE
    Run the script and follow the prompts:
    .\MSI_2_Intunewin.ps1

.CREATED BY: Mike Croucher

.DATE: 04/15/2025
#>

# Prompt user for the MSI folder location
$sourceFolder = Read-Host "Enter the full path to the folder containing the MSI files"

# Prompt user for the Intunewin output folder location
$outputFolder = Read-Host "Enter the full path to the folder where Intunewin files should be saved"

# Prompt user for the location of IntuneWinAppUtil.exe
$intuneWinAppUtilPath = Read-Host "Enter the full path to IntuneWinAppUtil.exe"

# Initialize log file
$logFile = Join-Path -Path $outputFolder -ChildPath "MSI_Processing_Log.txt"

# Function to log messages to console and log file
function Log-Message {
    param (
        [string]$Message,
        [string]$Color = "White"
    )
    Write-Host $Message -ForegroundColor $Color
    Add-Content -Path $logFile -Value "$((Get-Date).ToString()): $Message"
}

# Validate paths
if (-Not (Test-Path -Path $sourceFolder)) {
    Log-Message "The source folder '$sourceFolder' does not exist. Please create it and rerun the script." "Red"
    exit
}
if (-Not (Test-Path -Path $intuneWinAppUtilPath -PathType Leaf) -or `
    -Not (Get-Command $intuneWinAppUtilPath -ErrorAction SilentlyContinue)) {
    Log-Message "The IntuneWinAppUtil.exe file is not executable or not found at '$intuneWinAppUtilPath'." "Red"
    exit
}
if (-Not (Test-Path -Path $outputFolder)) {
    $confirm = Read-Host "The output folder '$outputFolder' does not exist. Do you want to create it? (Y/N)"
    if ($confirm -ne 'Y') {
        Log-Message "Output folder creation canceled. Exiting script." "Yellow"
        exit
    }
    Log-Message "Creating output folder..."
    New-Item -ItemType Directory -Path $outputFolder | Out-Null
}

# Start logging
Add-Content -Path $logFile -Value "Processing started at $(Get-Date)"

# Get all MSI files from the source folder
$msiFiles = Get-ChildItem -Path $sourceFolder -Filter *.msi
if ($msiFiles.Count -eq 0) {
    Log-Message "No MSI files found in the source folder '$sourceFolder'. Please add files and rerun the script." "Red"
    exit
}

# Initialize counters
$successCount = 0
$failureCount = 0

# Process each MSI file
$totalFiles = $msiFiles.Count
$currentFile = 0

foreach ($msiFile in $msiFiles) {
    $currentFile++
    Write-Progress -Activity "Processing MSI files" `
        -Status "Processing $currentFile of $totalFiles: $($msiFile.Name)" `
        -PercentComplete (($currentFile / $totalFiles) * 100)

    $msiFilePath = $msiFile.FullName
    $msiFileName = $msiFile.Name

    # Run the Win32 Content Prep Tool
    Log-Message "Processing $msiFileName..." "Yellow"
    try {
        Start-Process -FilePath $intuneWinAppUtilPath `
            -ArgumentList "-c `"$sourceFolder`" -s `"$msiFilePath`" -o `"$outputFolder`"" -Wait -NoNewWindow
        Log-Message "Successfully processed $msiFileName." "Green"
        $successCount++
    } catch {
        Log-Message "Failed to process $msiFileName. Error: $($_.Exception.Message)" "Red"
        $failureCount++
    }
}

# Display summary
Log-Message "Processing complete. Success: $successCount, Failed: $failureCount." "Cyan"
Write-Host "All MSI files have been processed and saved to $outputFolder."
Add-Content -Path $logFile -Value "Processing completed at $(Get-Date). Success: $successCount, Failed: $failureCount."