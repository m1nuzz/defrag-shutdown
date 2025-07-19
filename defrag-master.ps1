<#
.SYNOPSIS
Defragmentation Master Script.
Detects local fixed drives, prompts the user to select drives and shutdown option,
and saves the configuration into a separate execution script (defrag-now.ps1).

.DESCRIPTION
This script guides the user through selecting drives for defragmentation/optimization
and choosing whether to shut down the PC afterwards. It then generates a new
PowerShell script named 'defrag-now.ps1' containing these selected parameters
for direct execution later.

Includes fixes for:
- Correctly generating the boolean value for $DoShutdown in the execution script.
- More explicit string generation for the SelectedDrivesInfo array to avoid parsing errors.
- Updated optimization commands in the generated script based on user request.
  - SSD: Optimize-Volume -Analyze -ReTrim -Verbose
  - HDD: Optimize-Volume -Analyze -Defrag -Verbose
- Removed defrag.exe fallback and -Confirm from generated optimization commands.
- Added automatic execution of the generated defrag-now.ps1 script.
- Improved shutdown handling and error reporting in the generated script using shutdown.exe.
- Corrected debug message format in generated script.
- **Completely revised script generation using Add-Content and explicit string handling to ensure correct $DoShutdown value and debug output.**

.NOTES
Requires administrative privileges only for detecting media types accurately.
Saving the 'defrag-now.ps1' file typically does not require admin rights
unless saving to a protected location.
The generated 'defrag-now.ps1' requires administrative privileges to run Optimize-Volume or shutdown.exe.
**This master script should be run as administrator to automatically launch the generated script with necessary permissions.**
#>

# --- Configuration ---
$ExecutionScriptFile = Join-Path $PSScriptRoot "defrag-now.ps1" # The script to be generated

# --- Function to detect local fixed drives and their media type ---
function Get-LocalFixedDrivesWithMediaType {
    Write-Host "Detecting available local fixed drives and their media types..." -ForegroundColor Cyan
    $availableDrivesInfo = @()

    # Get all volumes that have a drive letter and are not removable or CD-ROM
    try {
        $volumes = Get-Volume | Where-Object {
            $null -ne $_.DriveLetter -and
            $_.DriveType -eq 'Fixed' -and
            $null -ne $_.FileSystem # Exclude volumes without a file system like System Reserved without a letter
        }

        foreach ($volume in $volumes) {
            $driveLetter = $volume.DriveLetter
            Write-Host "`n=== Processing drive letter $($driveLetter): ===" -ForegroundColor DarkCyan
            Write-Host "[DEBUG] Found volume $($driveLetter): with DriveType $($volume.DriveType)." -ForegroundColor DarkGray

            $MediaType = "Unknown"
            $ErrorOccurred = $false
            $diskNumber = $null # Initialize diskNumber

            try {
                # Get Partition object for the drive letter to find the Disk Number
                # Use ErrorAction Stop to catch issues with Get-Partition
                $partition = Get-Partition -DriveLetter $driveLetter -ErrorAction Stop
                $diskNumber = $partition.DiskNumber
                Write-Host "[DEBUG] Got DiskNumber $diskNumber for drive $($driveLetter):" -ForegroundColor DarkGray

                # --- Attempt 1: Get Media Type using Get-CimInstance MSFT_PhysicalDisk ---
                $cimPhysicalDisk = $null
                try {
                    # MSFT_PhysicalDisk DeviceId corresponds to the DiskNumber
                    $cimPhysicalDisk = Get-CimInstance -ClassName MSFT_PhysicalDisk -Namespace root\Microsoft\Windows\Storage | Where-Object { $_.DeviceId -eq $diskNumber } | Select-Object -First 1
                    if ($cimPhysicalDisk) {
                        # MSFT_PhysicalDisk MediaType values: 3=HDD, 4=SSD, etc. OR strings 'SSD', 'HDD'
                        switch ($cimPhysicalDisk.MediaType) {
                            3 { $MediaType = "HDD" }
                            4 { $MediaType = "SSD" }
                            "SSD" { $MediaType = "SSD" } # Correctly map string "SSD"
                            "HDD" { $MediaType = "HDD" } # Correctly map string "HDD"
                            default { $MediaType = "Unknown" }
                        }
                        Write-Host "[DEBUG] Got MediaType '$($cimPhysicalDisk.MediaType)' mapped to '$MediaType' from MSFT_PhysicalDisk for physical disk $diskNumber." -ForegroundColor DarkGray
                    }
                    else {
                        Write-Warning "Could not find MSFT_PhysicalDisk instance for disk number $diskNumber. Attempting Get-PhysicalDisk fallback."
                        $ErrorOccurred = $true # Mark as error for this method, but continue to next
                    }
                }
                catch {
                    Write-Warning "Error querying MSFT_PhysicalDisk for drive $($driveLetter): $($_.Exception.Message). Attempting Get-PhysicalDisk fallback."
                    $ErrorOccurred = $true # Mark as error, continue to next
                }

                # --- Attempt 2: Fallback to Get-PhysicalDisk if MSFT_PhysicalDisk failed ---
                if ($MediaType -eq "Unknown" -and $ErrorOccurred) {
                    $ErrorOccurred = $false # Reset error flag for this attempt
                    $physicalDisk = $null
                    try {
                        # Get-PhysicalDisk by DiskNumber (using -DiskNumber or -Number)
                        # Note: Depending on PowerShell version, -Number or -DiskNumber might be needed.
                        # Let's try -DiskNumber first as it's more common with Get-PhysicalDisk
                        $physicalDisk = Get-PhysicalDisk -DiskNumber $diskNumber -ErrorAction Stop
                        if ($physicalDisk) {
                            $MediaType = $physicalDisk.MediaType
                            Write-Host "[DEBUG] Got MediaType '$MediaType' from Get-PhysicalDisk for physical disk $diskNumber." -ForegroundColor DarkGray
                        }
                        else {
                            Write-Warning "Could not find matching PhysicalDisk object by DiskNumber $diskNumber. Attempting Win32_DiskDrive fallback."
                            $ErrorOccurred = $true # Mark as error, continue to next
                        }
                    }
                    catch {
                        # If -Number failed, try -Number as a last resort for Get-PhysicalDisk
                        try {
                            $physicalDisk = Get-PhysicalDisk -Number $diskNumber -ErrorAction Stop
                            if ($physicalDisk) {
                                $MediaType = $physicalDisk.MediaType
                                Write-Host "[DEBUG] Got MediaType '$MediaType' from Get-PhysicalDisk (using -Number) for physical disk $diskNumber." -ForegroundColor DarkGray
                            }
                            else {
                                Write-Warning "Could not find matching PhysicalDisk object by Number $diskNumber. Attempting Win32_DiskDrive fallback."
                                $ErrorOccurred = $true # Mark as error, continue to next
                            }
                        }
                        catch {
                            Write-Warning "Error querying Get-PhysicalDisk (both -DiskNumber and -Number failed) for drive $($driveLetter): $($_.Exception.Message). Attempting Win32_DiskDrive fallback."
                            $ErrorOccurred = $true # Mark as error, continue to next
                        }
                    }
                }

                # --- Attempt 3: Fallback to WMI (Win32_DiskDrive) if previous methods failed ---
                if ($MediaType -eq "Unknown" -and $ErrorOccurred) {
                    $ErrorOccurred = $false # Reset error flag for this attempt
                    try {
                        $wmiDisk = Get-CimInstance -ClassName Win32_DiskDrive | Where-Object { $_.DeviceID -eq "\\.\PHYSICALDRIVE$diskNumber" } | Select-Object -First 1

                        if ($wmiDisk) {
                            # WMI MediaType values: 0=Unknown, 1=Not Available, 3=HDD, 4=SSD, etc.
                            # Also check for string values like 'Fixed hard disk media' if numerical mapping fails
                            switch ($wmiDisk.MediaType) {
                                3 { $MediaType = "HDD" }
                                4 { $MediaType = "SSD" }
                                "SSD" { $MediaType = "SSD" } # Correctly map string "SSD"
                                "HDD" { $MediaType = "HDD" } # Correctly map string "HDD"
                                "Fixed hard disk media" {
                                    # This WMI value is ambiguous, cannot reliably determine SSD vs HDD
                                    $MediaType = "Unknown"
                                    Write-Warning "WMI MediaType 'Fixed hard disk media' is ambiguous. Cannot reliably determine SSD vs HDD."
                                }
                                default { $MediaType = "Unknown" }
                            }
                            Write-Host "[DEBUG] Got WMI MediaType '$($wmiDisk.MediaType)' mapped to '$MediaType' from Win32_DiskDrive for physical disk $diskNumber." -ForegroundColor DarkGray
                            # If WMI successfully found a type, clear the error flag for this drive
                            if ($MediaType -ne "Unknown") { $ErrorOccurred = $false }
                        }
                        else {
                            Write-Warning "Could not find WMI Win32_DiskDrive instance for disk number $diskNumber."
                            $ErrorOccurred = $true # Still an error if WMI also fails
                        }
                    }
                    catch {
                        Write-Warning "Error during Win32_DiskDrive WMI fallback for drive $($driveLetter): $($_.Exception.Message)"
                        $ErrorOccurred = $true # Ensure error is true if WMI fallback fails
                        $MediaType = "Unknown" # Ensure MediaType is Unknown on error
                    }
                }


            }
            catch {
                Write-Warning "Error getting disk information for drive $($driveLetter): $($_.Exception.Message)"
                $ErrorOccurred = $true
                $MediaType = "Unknown" # Ensure MediaType is Unknown on error
            }

            # Add drive info to the list
            $availableDrivesInfo += [PSCustomObject]@{
                DriveLetter = $driveLetter
                MediaType   = $MediaType
                IsValid     = ($MediaType -ne "Unknown") # Mark as valid ONLY if media type was determined
            }

            if ($MediaType -eq "Unknown") {
                Write-Host "Found local fixed drive: $($driveLetter): (Media Type: Could not determine or Unknown)" -ForegroundColor Yellow
            }
            else {
                Write-Host "Found local fixed drive: $($driveLetter): (Media Type: $MediaType)" -ForegroundColor Green
            }
            Write-Host "=== Finished checking drive letter $($driveLetter): ===" -ForegroundColor DarkCyan
        }
    }
    catch {
        Write-Error "An error occurred while detecting drives: $($_.Exception.Message)"
        # Continue script execution, but availableDrivesInfo might be empty
    }

    Write-Host "`n========================================" -ForegroundColor Cyan
    if ($availableDrivesInfo.Count -gt 0) {
        $driveSummary = $availableDrivesInfo | ForEach-Object { "$($_.DriveLetter): ($($_.MediaType))" } | Out-String
        Write-Host "FINAL Check: Available drives found:" -ForegroundColor Cyan
        Write-Host $driveSummary.Trim() -ForegroundColor Cyan
    }
    else {
        Write-Host "FINAL Check: No available drives found." -ForegroundColor Yellow
    }
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""

    return $availableDrivesInfo
}

# --- Main script logic ---

# Detect drives and their types
$availableDrivesInfo = Get-LocalFixedDrivesWithMediaType

# Filter for drives where we successfully got info AND the media type is not Unknown
$validAvailableDrivesInfo = $availableDrivesInfo | Where-Object { $_.IsValid }

if ($validAvailableDrivesInfo.Count -eq 0) {
    Write-Host "No valid local fixed drives found or could not reliably determine media type for any. Nothing to configure." -ForegroundColor Red
    Read-Host "Press Enter to exit."
    exit
}

# Display available drives with types before input
Write-Host "Available local fixed drives (Media Type Determined):" -ForegroundColor Cyan
$validAvailableDrivesInfo | Format-Table -AutoSize

# --- User input for drives ---
$userInput = Read-Host "Enter drive letters to include in the execution script (e.g., C D E), or ALL for all listed valid drives"
$userInput = $userInput.Trim().Replace(',', ' ') # Trim whitespace and replace commas with spaces

# --- Process input ---
$selectedDrivesInfo = @()

if ($userInput -eq "ALL") {
    $selectedDrivesInfo = $validAvailableDrivesInfo
}
else {
    $inputDriveLetters = $userInput.Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries) | ForEach-Object { $_.ToUpper() }

    foreach ($inputDrive in $inputDriveLetters) {
        # Find the drive info from the list of valid drives
        $foundDrive = $validAvailableDrivesInfo | Where-Object { $_.DriveLetter -eq $inputDrive } | Select-Object -First 1

        if ($foundDrive) {
            $selectedDrivesInfo += $foundDrive
        }
        else {
            Write-Warning "Drive $inputDrive is not a valid local fixed drive listed above or was not entered correctly. Skipping."
        }
    }
}

# Remove duplicates based on DriveLetter and sort
$selectedDrivesInfo = $selectedDrivesInfo | Sort-Object DriveLetter -Unique

if ($selectedDrivesInfo.Count -eq 0) {
    Write-Host "No valid local fixed drives selected to include in the execution script." -ForegroundColor Red
    Read-Host "Press Enter to exit."
    exit
}

Write-Host "You selected to include:" -ForegroundColor Green
$selectedDrivesInfo | Format-Table -AutoSize
Write-Host ""

# --- Question about shutdown --- [FIXED SECTION]
$doShutdown = $false
$shutdownChoice = Read-Host "Include option to Shutdown PC after defragmentation? (Y/N)"
$shutdownChoice = $shutdownChoice.Trim().ToUpper() # Convert to uppercase to be more forgiving

if ($shutdownChoice -eq "Y") {
    # Case-insensitive check for 'Y'
    $doShutdown = $true
    Write-Host "Shutdown option will be included in the execution script." -ForegroundColor Yellow
}
else {
    $doShutdown = $false # Explicitly set to false
    Write-Host "PC will remain on after defragmentation." -ForegroundColor Yellow
}
Write-Host ""

# --- Generate and Save Execution Script ---
Write-Host "Generating execution script '$ExecutionScriptFile' using Add-Content..." -ForegroundColor Cyan

# --- Build the script content using Add-Content ---

# Add header comments and parameters section start
$headerContent = @"
<#
.SYNOPSIS
Defragmentation Execution Script.
Performs defragmentation/optimization based on pre-configured drives and shutdown option.

.DESCRIPTION
This script is generated by defrag-master.ps1. It contains hardcoded parameters
for the drives to optimize and whether to shut down the PC afterwards.
It requires administrative privileges to run Optimize-Volume or shutdown.exe.

.NOTES
Generated by defrag-master.ps1.
Requires administrative privileges.
Uses Optimize-Volume with Analyze, Verbose, and appropriate operation (ReTrim for SSD, Defrag for HDD).
Includes improved shutdown handling and logging using shutdown.exe.
#>

# --- Parameters (Hardcoded from Master Script) ---
"@
$headerContent | Set-Content -Path $ExecutionScriptFile -Force -Encoding UTF8

# Add the SelectedDrivesInfo array definition
# Build the string representation of the SelectedDrivesInfo array elements explicitly
$selectedDrivesInfoLines = @()
$selectedDrivesInfoLines += "`$SelectedDrivesInfo = @(" # Add opening line
foreach ($driveInfo in $selectedDrivesInfo) {
    # Add each PSCustomObject definition as a line, with comma unless it's the last one
    $line = "[PSCustomObject]@{ DriveLetter = '$($driveInfo.DriveLetter)'; MediaType = '$($driveInfo.MediaType)' }"
    # Add comma and newline if it's not the last element
    if ($driveInfo -ne $selectedDrivesInfo[-1]) {
        $line += ","
    }
    $selectedDrivesInfoLines += $line
}
$selectedDrivesInfoLines += ")" # Add closing line
$selectedDrivesInfoString = $selectedDrivesInfoLines -Join "`n"
Add-Content -Path $ExecutionScriptFile -Value $selectedDrivesInfoString -Encoding UTF8

# Add the $DoShutdown parameter line - FIXED VERSION
# This ensures we write the literal strings "$true" or "$false" to the file
if ($doShutdown) {
    Add-Content -Path $ExecutionScriptFile -Value '$DoShutdown = $true' -Encoding UTF8
}
else {
    Add-Content -Path $ExecutionScriptFile -Value '$DoShutdown = $false' -Encoding UTF8
}

# Add debug lines for $DoShutdown value and type
Add-Content -Path $ExecutionScriptFile -Value "`n# --- Debugging: Show the value and type of `$DoShutdown ---" -Encoding UTF8
# Using single quotes to ensure literal strings are written to the file
Add-Content -Path $ExecutionScriptFile -Value 'Write-Host "[DEBUG] $DoShutdown Value:" -NoNewline -ForegroundColor DarkGray' -Encoding UTF8
Add-Content -Path $ExecutionScriptFile -Value 'Write-Host $DoShutdown -ForegroundColor DarkGray' -Encoding UTF8
Add-Content -Path $ExecutionScriptFile -Value '$DoShutdownType = $DoShutdown.GetType().Name' -Encoding UTF8
Add-Content -Path $ExecutionScriptFile -Value 'Write-Host "[DEBUG] $DoShutdown Type:" -NoNewline -ForegroundColor DarkGray' -Encoding UTF8
Add-Content -Path $ExecutionScriptFile -Value 'Write-Host $DoShutdownType -ForegroundColor DarkGray' -Encoding UTF8


# Add the rest of the script content
$footerContent = @"

# --- Check for Administrative Privileges ---
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Error "This script requires administrative privileges to run Optimize-Volume or shutdown.exe."
    Write-Host "Please right-click the script file and select 'Run as administrator'." -ForegroundColor Red
    Read-Host "Press Enter to exit."
    exit 1
}

Write-Host "Starting defragmentation process at `$(Get-Date -Format 'HH:mm:ss on dd-MM-yyyy')" -ForegroundColor Cyan
Write-Host ""

# --- Defragmentation Process ---
foreach (`$driveInfo in `$SelectedDrivesInfo) {
    `$driveLetter = `$driveInfo.DriveLetter
    `$MediaType = `$driveInfo.MediaType

    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host "Processing drive `$(`$driveLetter):" -ForegroundColor Cyan
    Write-Host "[DEBUG] Media Type determined earlier: `$MediaType" -ForegroundColor DarkGray
    Write-Host ""

    # Select optimization command based on media type
    switch (`$MediaType) {
        "SSD" {
            Write-Host "Drive `$(`$driveLetter): identified as SSD (`$MediaType). Performing optimization (Optimize-Volume -Analyze -ReTrim -Verbose)..." -ForegroundColor Green
            try {
                Optimize-Volume -DriveLetter `$driveLetter -Analyze -ReTrim -Verbose -ErrorAction Stop
            } catch {
                 Write-Error "Error performing SSD optimization on drive `$(`$driveLetter): `$(`$_.Exception.Message)"
            }
        }
        "HDD" {
            Write-Host "Drive `$(`$driveLetter): identified as HDD (`$MediaType). Performing optimization (Optimize-Volume -Analyze -Defrag -Verbose)..." -ForegroundColor Green
             try {
                Optimize-Volume -DriveLetter `$driveLetter -Analyze -Defrag -Verbose -ErrorAction Stop
            } catch {
                 Write-Error "Error performing HDD optimization on drive `$(`$driveLetter): `$(`$_.Exception.Message)"
            }
        }
        default {
            # This case should ideally not be reached if only valid drives are included,
            # but included as a fallback
            Write-Host "Drive `$(`$driveLetter): media type is Unknown/Other (`$MediaType). Performing standard optimization (Optimize-Volume -Analyze -Defrag -Verbose)..." -ForegroundColor Yellow
             try {
                # Analyze first, then Defrag for unknown types
                Optimize-Volume -DriveLetter `$driveLetter -Analyze -Defrag -Verbose -ErrorAction Stop
            } catch {
                 Write-Error "Error performing standard optimization on drive `$(`$driveLetter): `$(`$_.Exception.Message)"
            }
        }
    }

    Write-Host "`nFinished processing drive `$(`$driveLetter):" -ForegroundColor Cyan
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host ""
}

# --- Shutdown if selected ---
if (`$DoShutdown) {
    Write-Host "`nDefragmentation tasks complete. Attempting to shut down in 10 seconds using shutdown.exe..." -ForegroundColor Yellow
    # Use shutdown.exe for potentially more reliable shutdown from script
    # Wrap in try/catch to report shutdown errors
    try {
        Write-Host "[DEBUG] \$DoShutdown is `$DoShutdown. Initiating shutdown via shutdown.exe..." -ForegroundColor DarkGray
        # Use & to run external command
        & shutdown.exe /s /t 10 /f
        Write-Host "[DEBUG] shutdown.exe command executed. Script execution may stop now." -ForegroundColor DarkGray
        # Note: Script execution might stop here if shutdown is successful.
    } catch {
        Write-Error "Failed to initiate shutdown via shutdown.exe: `$(`$_.Exception.Message)"
        Write-Host "Please shut down your computer manually." -ForegroundColor Red
        Read-Host "Press Enter to exit." # Keep window open if shutdown fails
    }
} else {
    Write-Host "`nDefragmentation tasks complete at `$(Get-Date -Format 'HH:mm:ss on dd-MM-yyyy')." -ForegroundColor Green
    Write-Host ""
    Write-Host "Script finished." -ForegroundColor Cyan
    Read-Host "Press Enter to exit."
}

"@ # End of footer Here-String
Add-Content -Path $ExecutionScriptFile -Value $footerContent -Encoding UTF8


Write-Host "Execution script '$ExecutionScriptFile' generated successfully." -ForegroundColor Green
Write-Host "Launching '$ExecutionScriptFile'..." -ForegroundColor Cyan

# --- Launch the generated script ---
# Use Start-Process to run the script in a new window, ensuring it runs as administrator
# if the master script was run as administrator.
Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile", "-ExecutionPolicy Bypass", "-File", "`"$ExecutionScriptFile`"" -Verb RunAs
