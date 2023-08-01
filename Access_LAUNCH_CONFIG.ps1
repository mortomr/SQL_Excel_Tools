
# Modify this to take PDDR offline
$fileName = "PDS.accde"
# fileName = "PDS Offline.accde"
$sourcePath = "\\jdc-fas3270-a1\Nuclear Project Delivery\Database\PEI Dashboard\$fileName"
$destinationPath = "C:\Temp"
$logfile = Join-Path -Path $env:USERPROFILE -ChildPath "PDDR.log"

# Function to log output to a file
function Write-Log {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Message,

        [Parameter(Mandatory=$true)]
        [string]$LogFile
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - $Message"

    Add-Content -Path $LogFile -Value $logMessage
}

# Function to test network connectivity and resolve IP address
function Test-NetworkConnection {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ComputerName,

        [Parameter(Mandatory=$true)]
        [string]$LogFile
    )

    try {
        $pingResult = Test-Connection -ComputerName $ComputerName -Count 4

        # Extract the average round-trip time from the ping result
        $averageResponseTime = ($pingResult | Measure-Object -Property ResponseTime -Average).Average

        $pingInfo = "Ping to $ComputerName successful.`n"
        $pingInfo += "Average Response Time: $averageResponseTime ms`n"

        Write-Log -Message $pingInfo -LogFile $LogFile
    }
    catch {
        $errorMessage = "Ping to $ComputerName failed. Error: $($_.Exception.Message)`n"
        Write-Log -Message $errorMessage -LogFile $LogFile
    }
}

# Function to terminate MS Access processes
function Terminate-AccessProcesses {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ProcessName,

        [Parameter(Mandatory=$true)]
        [string]$LogFile
    )

    $accessProcesses = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue

    if ($accessProcesses) {
        foreach ($process in $accessProcesses) {
            $process | Stop-Process -Force
        }

        $accessProcessInfo = "MS Access processes terminated.`n"
        Write-Log -Message $accessProcessInfo -LogFile $LogFile
    }
    else {
        $accessProcessInfo = "No running MS Access processes found."
        Write-Log -Message $accessProcessInfo -LogFile $LogFile
    }
}

# Function to test destination folder files and delete them
function Test-AndDeleteDestinationFolderFiles {
    param (
        [Parameter(Mandatory=$true)]
        [string]$destinationPath,

        [Parameter(Mandatory=$true)]
        [string]$LogFile
    )

    try {
        $allFiles = Get-ChildItem -Path $destinationPath -File -Force

        foreach ($file in $allFiles) {
            if ($file.Extension -eq ".mdb" -or $file.Extension -eq ".accdb" -or $file.Extension -eq ".accde" -or $file.Extension -eq ".laccdb") {
                $fileName = $file.Name
                $lastWriteDate = $file.LastWriteTime
                $fileSize = $file.Length

                $fileInfo = "File Name: $fileName`n"
                $fileInfo += "Last Write Date: $lastWriteDate`n"
                $fileInfo += "File Size: $fileSize bytes`n"

                # Check file permissions
                try {
                    $acl = Get-Acl -Path $file.FullName
                    $fileInfo += "User Permissions:`n"
                    foreach ($accessRule in $acl.Access) {
                        $fileInfo += "- $($accessRule.IdentityReference): $($accessRule.FileSystemRights)`n"
                    }
                }
                catch {
                    $fileInfo += "Error retrieving file permissions: $($_.Exception.Message)`n"
                }

                # Delete the file
                try {
                    Remove-Item -Path $file.FullName -Force
                    $fileInfo += "File deleted successfully`n"
                }
                catch {
                    $fileInfo += "Error deleting the file: $($_.Exception.Message)`n"
                }

                $fileInfo += "--------------------------------`n"
                Write-Log -Message $fileInfo -LogFile $LogFile
            }
        }
    }
    catch {
        $errorMessage = "Error retrieving files from the destination folder: $($_.Exception.Message)"
        Write-Log -Message $errorMessage -LogFile $LogFile
    }
}

# Function to copy and launch a file
function Copy-AndLaunchFile {
    param (
        [Parameter(Mandatory=$true)]
        [string]$SourcePath,

        [Parameter(Mandatory=$true)]
        [string]$DestinationPath,

        [Parameter(Mandatory=$true)]
        [string]$FileName
    )

    try {
        # Copy the file to the destination folder
        $destinationFilePath = Join-Path -Path $DestinationPath -ChildPath "MyLocalPDDR.accde" #$FileName
        Copy-Item -Path $SourcePath -Destination $destinationFilePath -Force

        # Unblock the copied file
        Unblock-File -Path $destinationFilePath

        # Launch the copied file
        Start-Process -FilePath $destinationFilePath -ErrorAction Stop

        $successMessage = "File '$FileName' copied and launched successfully."
        Write-Host $successMessage
    }
    catch {
        $errorMessage = "Error copying or launching the file '$FileName': $($_.Exception.Message)"
        Write-Host $errorMessage
    }
}


# Main script

Write-Log -Message "+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ New Session." -LogFile $logfile
Write-Log -Message "PDDR Launch Testing underway." -LogFile $logfile
Write-Log -Message "Testing Memory" -LogFile $logfile

# Get system memory information
$systemMemory = Get-CimInstance -Class Win32_ComputerSystem | Select-Object TotalPhysicalMemory, FreePhysicalMemory

# Calculate used and available memory
$usedMemory = $systemMemory.TotalPhysicalMemory - $systemMemory.FreePhysicalMemory
$availableMemory = $systemMemory.FreePhysicalMemory

# Format memory sizes to be more readable
$usedMemoryFormatted = "{0:N2} GB" -f ($usedMemory / 1GB)
$availableMemoryFormatted = "{0:N2} GB" -f ($availableMemory / 1GB)

Write-Log -Message "Testing Network" -LogFile $logfile

# Test network connectivity and resolve IP address
$computerName = 'JDCDBETSP682'  # Replace with the computer name you want to resolve
$ipAddress = (Resolve-DnsName -Name $computerName -ErrorAction SilentlyContinue).IPAddress

if ($ipAddress) {
    Write-Log -Message "IP address for computer '$computerName' is: $ipAddress" -LogFile $logfile
    Test-NetworkConnection -ComputerName $ipAddress -LogFile $logfile
} else {
    Write-Log -Message "Failed to resolve the IP address for computer '$computerName'." -LogFile $logfile
}

Write-Log -Message "Testing Open Access Processes" -LogFile $logfile

# Specify the name of the MS Access process
$processName = "MSACCESS"

# Terminate MS Access processes
Terminate-AccessProcesses -ProcessName $processName -LogFile $logfile

Write-Log -Message "Testing Destination Folder" -LogFile $logfile

# Test destination folder files
# Test-DestinationFolderFiles -FolderPath $destinationPath -LogFile $logfile
# Log destination folder files
Test-AndDeleteDestinationFolderFiles -FolderPath $destinationPath -LogFile $logFile

Write-Log -Message "Download and Launch Processes Started" -LogFile $logfile
# Call the function to copy and launch the file
Copy-AndLaunchFile -SourcePath $sourcePath -DestinationPath $destinationPath -FileName $fileName

Write-Log -Message "System Memory:`nUsed Memory: $usedMemoryFormatted`nAvailable Memory: $availableMemoryFormatted" -LogFile $logfile

# Get system uptime
$uptime = (Get-CimInstance -Class Win32_OperatingSystem).LastBootUpTime
$uptime = (Get-Date) - $uptime
Write-Log -Message "System Uptime: $uptime" -LogFile $logfile


Write-Log -Message "File information exported to $logfile" -LogFile $logfile



# Output any issues encountered during download or launch, along with troubleshooting notes
if (Test-Path -Path $logfile) {
    Write-Host "Issues encountered. Check the PDDR_Test_Launch.log file for details."
}

# Prompt for user confirmation before exiting
# Read-Host "Press Enter to exit"
