# Define the paths
$msiFolder = "C:\Users\croucherm\Desktop\Intune\MSI"
$outputFolder = "C:\Users\croucherm\Desktop\Intune\Intunewin"
$processedFolder = "C:\Users\croucherm\Desktop\Intune\ProcessedMSIs"
$util = "C:\Users\croucherm\Desktop\Intune\IntuneWinAppUtil.exe"
$logFile = "C:\Users\croucherm\Desktop\Intune\IntunePackagingLog.txt"

# Create required folders if they don't exist
foreach ($path in @($msiFolder, $outputFolder, $processedFolder)) {
    if (-not (Test-Path $path)) {
        New-Item -Path $path -ItemType Directory | Out-Null
    }
}

# Logging function
function Write-Log {
    param ([string]$message)
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $fullMessage = "$timestamp - $message"
    $fullMessage | Out-File -Append -FilePath $logFile
    Write-Host $fullMessage
}

# Get all MSI files
$msiFiles = Get-ChildItem -Path $msiFolder -Filter *.msi
$total = $msiFiles.Count
$successCount = 0
$failureCount = 0

Write-Log "========== Intune Win32 Packaging Started =========="
Write-Log "Total files found: $total"
Write-Log "Tool path: $util"
Write-Log "Output folder: $outputFolder"
Write-Log "----------------------------------------------------"

# Loop through MSI files
for ($i = 0; $i -lt $total; $i++) {
    $msiFile = $msiFiles[$i]
    $fileNumber = $i + 1
    Write-Log "[$fileNumber of $total] Processing: $($msiFile.Name)"

    $outputFile = Join-Path $outputFolder ([System.IO.Path]::GetFileNameWithoutExtension($msiFile) + ".intunewin")
    $args = "-c `"$msiFolder`" -s `"$($msiFile.Name)`" -o `"$outputFolder`""

    try {
        Write-Log "Running: $util $args"
        $process = Start-Process -FilePath $util -ArgumentList $args -Wait -PassThru -NoNewWindow

        if ($process.ExitCode -eq 0) {
            if (Test-Path $outputFile) {
                $size = (Get-Item $outputFile).Length
                Write-Log "Success: $($msiFile.Name) → $([Math]::Round($size / 1MB, 2)) MB"
                Move-Item -Path $msiFile.FullName -Destination $processedFolder
                Write-Log "Moved to: $processedFolder"
                $successCount++
            } else {
                Write-Log "Warning: No output created for $($msiFile.Name)"
                $failureCount++
            }
        } else {
            Write-Log "Error: Exit code $($process.ExitCode) for $($msiFile.Name)"
            $failureCount++
        }
    }
    catch {
        Write-Log "Exception while processing $($msiFile.Name): $($_.Exception.Message)"
        $failureCount++
    }

    Write-Log "----------------------------------------------------"
}

Write-Log "All done!"
Write-Log "Successes: $successCount"
Write-Log "Failures: $failureCount"
Write-Log "========== Intune Win32 Packaging Complete =========="
