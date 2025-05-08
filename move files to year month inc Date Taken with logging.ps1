<#
This script organizes files in a directory. 
1. It detects duplicates using SHA256 hashes and moves them to a "duplicates" folder.
2. It organizes the remaining files by their date:
   - For images: uses the "Date Taken" metadata if available
   - For other files: uses their last write date
   Files are organized into Year/Month subfolders.
The script includes comprehensive logging and validation for directory creation.

Configure the `$Data2SortPath` variable to the directory to process.
#>

param (
    [string]$Data2SortPath = "C:\Desktops",
    [string]$duplicatesPath = "$Data2SortPath\duplicates"
)

# Define log file path
$LogFile = "$($Data2SortPath)\OrganizeFilesLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

# Function to log messages
function Write-MoveLog {
    param (
        [string]$Message,
        [string]$LogLevel = "INFO"
    )
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$Timestamp [$LogLevel] $Message" | Out-File -FilePath $LogFile -Append
}

$MoveLogOutput = "$($Data2SortPath)\OrganizeFilesTranscript_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

# Log script start
Write-MoveLog "Script started. Organizing files in '$Data2SortPath'." "INFO"

# Validate input directory
if (-not (Test-Path -Path $Data2SortPath)) {
    Write-MoveLog "Error: The specified path '$Data2SortPath' does not exist." "ERROR"
    throw "The specified path does not exist."
}

# Create duplicates folder with validation
try {
    if (-not (Test-Path -Path $duplicatesPath)) {
        New-Item -Path $duplicatesPath -ItemType Directory -Force | Out-Null
        Write-MoveLog "Created duplicates folder at '$duplicatesPath'." "INFO"
    } else {
        Write-MoveLog "Duplicates folder already exists at '$duplicatesPath'." "INFO"
    }
} catch {
    Write-MoveLog "Failed to create duplicates folder: $_" "ERROR"
    throw $_
}

# Function to get date taken from image file if available
# Add necessary assembly for the image metadata extraction
Add-Type -AssemblyName System.Drawing

function Get-DateTaken {
    param (
        [string]$FilePath
    )
    
    try {
        # Check if this is an image file
        $imageExtensions = @('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.tif', '.mp4') #, '.*')  # can be any file
        $extension = [System.IO.Path]::GetExtension($FilePath).ToLower()
        
        if ($imageExtensions -contains $extension) {
            # Method 1: Using Shell.Application
            try {
                $shell = New-Object -ComObject Shell.Application
                $folder = $shell.Namespace([System.IO.Path]::GetDirectoryName($FilePath))
                $file = $folder.ParseName([System.IO.Path]::GetFileName($FilePath))
                
                # Try different property indices that might contain the date taken
                # Property indices can vary by Windows version
                foreach ($propertyIndex in @(12, 14, 208, 36867)) {
                    $dateTakenString = $folder.GetDetailsOf($file, $propertyIndex)
                    
                    if (![string]::IsNullOrWhiteSpace($dateTakenString)) {
                        try {
                            $dateObj = [DateTime]::Parse($dateTakenString)
                            Write-MoveLog "Found Date Taken via Shell.Application for '$FilePath': $dateObj (Property $propertyIndex)" "INFO"
                            
                            # Release COM objects
                            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($file) | Out-Null
                            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($folder) | Out-Null
                            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell) | Out-Null
                            
                            return $dateObj
                        } catch {
                            # Continue to next property or method if parsing fails
                        }
                    }
                }
                
                # Release COM objects if no date was found
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($file) | Out-Null
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($folder) | Out-Null
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell) | Out-Null
            } catch {
                Write-MoveLog "Error using Shell.Application for '$FilePath': $_" "WARNING"
            }
            
            # Method 2: Using System.Drawing for JPG/JPEG files
            if ($extension -eq '.jpg' -or $extension -eq '.jpeg') {
                try {
                    $image = [System.Drawing.Image]::FromFile($FilePath)
                    
                    # Get the PropertyItems
                    $propItems = $image.PropertyItems
                    
                    # Look for the date taken property (0x9003 = 36867)
                    $dateTakenPropertyItem = $propItems | Where-Object { $_.Id -eq 36867 }
                    
                    if ($dateTakenPropertyItem) {
                        $dateTakenString = [System.Text.Encoding]::ASCII.GetString($dateTakenPropertyItem.Value)
                        # EXIF date format is typically "YYYY:MM:DD HH:MM:SS"
                        $dateTakenString = $dateTakenString.Trim([char]0)  # Remove null terminators
                        
                        if ($dateTakenString -match '(\d{4}):(\d{2}):(\d{2}) (\d{2}):(\d{2}):(\d{2})') {
                            $dateObj = [DateTime]::new([int]$Matches[1], [int]$Matches[2], [int]$Matches[3], 
                                                      [int]$Matches[4], [int]$Matches[5], [int]$Matches[6])
                            
                            Write-MoveLog "Found Date Taken via EXIF for '$FilePath': $dateObj" "INFO"
                            $image.Dispose()
                            return $dateObj
                        }
                    }
                    
                    $image.Dispose()
                } catch {
                    Write-MoveLog "Error reading EXIF data for '$FilePath': $_" "WARNING"
                }
            }
        }
    } catch {
        Write-MoveLog "Error getting date taken for '$FilePath': $_" "WARNING"
    }
    
    Write-MoveLog "No date taken found for '$FilePath', will use LastWriteTime" "INFO"
    return $null
}

# Retrieve all files except those in the duplicates folder
$gtFiles = Get-ChildItem -Path $Data2SortPath -Recurse -File | Where-Object { 
    $_.DirectoryName -notmatch "duplicates" -and 
    $_.Extension -notmatch ".log" -and 
    $_.Extension -notmatch ".zip" 
}

# Hashtable for storing file hashes
$gtFileHashes = @{}

# Process each file
foreach ($file in $gtFiles) {
    try {
        # Calculate the file hash
        $gtFileHash = (Get-FileHash -Algorithm SHA256 -Path $file.FullName).Hash

        if ($gtFileHashes.ContainsKey($gtFileHash)) {
            # Handle duplicate file
            $duplicateName = Join-Path -Path $duplicatesPath -ChildPath $file.Name
            $counter = 1

            # Ensure the duplicate has a unique name
            while (Test-Path -Path $duplicateName) {
                $duplicateName = Join-Path $duplicatesPath -ChildPath ("{0}_{1}{2}" -f $file.BaseName, $counter, $file.Extension)
                $counter++
            }

            # Move duplicate to duplicates folder
            Move-Item -Path $file.FullName -Destination $duplicateName
            Write-MoveLog "Duplicate detected: '$($file.FullName)' moved to '$duplicateName'." "INFO"
        } 
        else {
            # Add unique file hash to the hashtable
            $gtFileHashes[$gtFileHash] = $file.FullName

            # Try to get date taken for images, fall back to LastWriteTime for other files
            $fileDate = Get-DateTaken -FilePath $file.FullName
            
            # If date taken is not available, use last write time
            if ($null -eq $fileDate) {
                $fileDate = $file.LastWriteTime
                Write-MoveLog "Using LastWriteTime for '$($file.FullName)': $fileDate" "INFO"
            } else {
                Write-MoveLog "Using DateTaken for '$($file.FullName)': $fileDate" "INFO"
            }

            $year = $fileDate.Year
            $monthName = (Get-Culture).DateTimeFormat.GetMonthName($fileDate.Month)

            # Create year and month directories with validation
            $yearPath = Join-Path -Path $Data2SortPath -ChildPath $year
            $monthPath = Join-Path -Path $yearPath -ChildPath $monthName

            try {
                if (-not (Test-Path -Path $yearPath)) {
                    New-Item -ItemType Directory -Path $yearPath -Force | Out-Null
                    Write-MoveLog "Created year folder at '$yearPath'." "INFO"
                }
                if (-not (Test-Path -Path $monthPath)) {
                    New-Item -ItemType Directory -Path $monthPath -Force | Out-Null
                    Write-MoveLog "Created month folder at '$monthPath'." "INFO"
                }
            } 
            catch {
                Write-MoveLog "Failed to create directory structure for '$($file.FullName)': $_" "ERROR"
                continue
            }

            # Move file to the appropriate folder
            $destination = Join-Path -Path $monthPath -ChildPath $file.Name
            $counter = 1

            try {
                # Check if file with same name already exists at destination
                while (Test-Path -Path $destination) {
                    $destination = Join-Path -Path $monthPath -ChildPath ("{0}_{1}{2}" -f $file.BaseName, $counter, $file.Extension)
                    $counter++
                }

                # Move the file
                Move-Item -Path $file.FullName -Destination $destination
                Write-MoveLog "File '$($file.FullName)' moved to '$destination'." "INFO"
            } 
            catch {
                Write-MoveLog "Failed to move file '$($file.FullName)': $_" "ERROR"
                continue
            }
        }
    } 
    catch {
        # Log and continue on error
        Write-MoveLog "Error processing file '$($file.FullName)': $_" "ERROR"
        continue
    }
}

# Clean up and force garbage collection
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

# Log script completion
Write-MoveLog "File organization complete. See details in the log file at '$LogFile'." "INFO"
Write-Host "Processing complete. Logs are saved to '$LogFile'."
