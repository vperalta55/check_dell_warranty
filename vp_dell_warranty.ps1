# This script installs Dell Command | Monitor, queries warranty information, and then uninstalls the software.
# It also cleans up the downloaded installer file.

# URL of the Dell Command | Monitor installer
$downloadUrl = "https://dl.dell.com/FOLDER09515677M/1/Dell-Command-Monitor_5RFFM_WIN_10.9.0.307_A00.EXE"
# Path to save the installer in the temporary directory
$downloadPath = "$env:TEMP\Dell-Command-Monitor_5RFFM_WIN_10.9.0.307_A00.EXE"
# Arguments for silent installation
$installArgs = "/s"

# Check if Dell Command | Monitor is already installed
$appInstalled = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "Dell Command | Monitor"} | Select-Object -First 1) -ne $null

if (!$appInstalled) {
    try {
        # Download the Dell Command | Monitor installer to the temp folder
        $webClient = New-Object System.Net.WebClient
        $webClient.DownloadFile($downloadUrl, $downloadPath)

        # Install Dell Command | Monitor silently
        Start-Process -FilePath $downloadPath -ArgumentList $installArgs -Wait
    } catch {
        Write-Host "An error occurred while downloading or installing Dell Command | Monitor: $_"
        exit 1
    }

    # Check if Dell Command | Monitor was successfully installed
    $appInstalled = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "Dell Command | Monitor"} | Select-Object -First 1) -ne $null
}

if ($appInstalled) {
    Write-Host "Dell Command | Monitor is installed."

    # Pause for 10 seconds to ensure the installation is complete
    Start-Sleep -Seconds 10
    
    try {
        # Retrieve warranty information using CIM
        $warrantyInfo = Get-CimInstance -Namespace root/dcim/sysman -ClassName DCIM_AssetWarrantyInformation
        if ($warrantyInfo) {
            $endDate = ($warrantyInfo | Sort-Object -Property WarrantyEndDate | Select -Last 1).WarrantyEndDate
            $startDate = ($warrantyInfo | Sort-Object -Property WarrantyStartDate | Select -Last 1).WarrantyStartDate

            Write-Host "End day: $endDate"
            Write-Host "Start day: $startDate"

            # Write the dates to text files
            Set-Content -Path "C:\temp\EndDay.txt" -Value $endDate
            Set-Content -Path "C:\temp\StartDay.txt" -Value $startDate
        } else {
            Write-Host "No warranty information found."
        }
    } catch {
        Write-Host "An error occurred while retrieving warranty information: $_"
    }

    try {
        # Uninstall Dell Command | Monitor
        $uninstallArgs = "/s /v/qn"
        $uninstallPath = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "Dell Command | Monitor"} | Select-Object -First 1).UninstallString
        if ($uninstallPath) {
            Start-Process -FilePath msiexec.exe -ArgumentList "$uninstallPath $uninstallArgs" -Wait
        } else {
            Write-Host "Uninstall path for Dell Command | Monitor not found."
        }
    } catch {
        Write-Host "An error occurred while uninstalling Dell Command | Monitor: $_"
    }
} else {
    Write-Host "Dell Command | Monitor installation failed."
}

# Clean up the Dell Command | Monitor installer from the temp folder
if (Test-Path $downloadPath) {
    Remove-Item -Path $downloadPath
}
