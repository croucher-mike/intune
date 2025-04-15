# Define paths
$sourceFolder = "C:\Users\croucherm\Desktop\Intune\MSI"  # Folder containing the .msi and other files
$setupFile = "YourInstaller.msi"                 # Name of the .msi file
$outputFolder = "C:\Users\croucherm\Desktop\Intune\Intunewin"       # Folder where the .intunewin file will be saved

# Run the Win32 Content Prep Tool
Start-Process -FilePath ".\IntuneWinAppUtil.exe" -ArgumentList "-c $sourceFolder -s $setupFile -o $outputFolder" -Wait
