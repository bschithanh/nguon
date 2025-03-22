# Check the instructions here on how to use it https://github.com/bschithanh

# Define variables
$tempDir = "$env:TEMP\RE0GA1NA-3MD1-AO3L-N3WO-5DT4EN5RE5TN5I"
$url = "https://raw.githubusercontent.com/bschithanh/nguon/main/ActivateAIO-ToolsV3.1.3.rar"
$output = "$tempDir\ActivateAIO-ToolsV3.1.3.rar"
$extractDir = "$tempDir"

# Ensure the temporary directory exists
if (!(Test-Path -Path $tempDir)) {
    New-Item -ItemType Directory -Path $tempDir | Out-Null
}

# Download the ZAR file to the temporary directory
Write-Host "Downloading the archive from $url"
Invoke-WebRequest -Uri $url -OutFile $output

# Extract the ZAR file
Write-Host "Extracting the necessary files for your system"
Expand-Archive -Path $output -DestinationPath $extractDir -Force

# Run the batch file and wait for it to finish
$batchFile = "$extractDir\ActivateAIO-ToolsV3.1.3\Activate AIO Tools v3.1.3 by Savio.cmd"
if (Test-Path -Path $batchFile) {
    Write-Host "Running the batch script..."
    Start-Process -FilePath "cmd.exe" -ArgumentList "/c `"$batchFile`"" -Wait
} else {
    Write-Host "Error: Batch script not found in $extractDir."
}

# Clean up: Delete the extracted directory after script.bat has finished
Write-Host "Cleaning up extracted files..."
Remove-Item -Path $output -Force
Remove-Item -Path $extractDir -Recurse -Force

# Clean up the main temp directory
Write-Host "Cleaning up main temp directory..."
Write-Host "Operation completed successfully."