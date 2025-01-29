<#
This script organizes files in a directory. 
1. It detects duplicates using SHA256 hashes and moves them to a "duplicates" folder.
2. It organizes the remaining files by their last write date into Year/Month subfolders.
The script includes comprehensive logging and validation for directory creation.

Configure the `$Data2SortPath` variable to the directory to process.
#>

param (
    [string]$Data2SortPath = "C:\Users\Admin\Desktop\Pics",
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

# Retrieve all files except those in the duplicates folder
$gtFiles = Get-ChildItem -Path $Data2SortPath -Recurse -File | Where { $_.DirectoryName -notmatch "duplicates" -or $_.extension -notmatch ".log" -and $_.extension -notmatch ".zip" }

# Hashtable for storing file hashes
$gtFileHashes = @{}

# Process each file
foreach ($file in $gtFiles) 
    {
        try 
            {
                # Calculate the file hash
                $gtFileHash = (Get-FileHash -Algorithm SHA256 -Path $file.FullName).Hash

                if ($gtFileHashes.ContainsKey($gtFileHash)) 
                    {
                        # Handle duplicate file
                        $duplicateName = Join-Path -Path $duplicatesPath -ChildPath $file.Name
                        $counter = 1

                        # Ensure the duplicate has a unique name
                        while (Test-Path -Path $duplicateName) 
                            {
                                $duplicateName = Join-Path $duplicatesPath -ChildPath ("{0}_{1}{2}" -f $file.BaseName, $counter, $file.Extension)
                                $counter++
                            }

                        # Move duplicate to duplicates folder
                        Move-Item -Path $file.FullName -Destination $duplicateName #-Verbose
                            Write-MoveLog "Duplicate detected: '$($file.FullName)' moved to '$duplicateName'." "INFO"
                    } 
                else 
                    {
                        # Add unique file hash to the hashtable
                        $gtFileHashes[$gtFileHash] = $file.FullName

                        # Organize file by date
                        $year = $file.LastWriteTime.Year
                        $monthName = (Get-Culture).DateTimeFormat.GetMonthName($file.LastWriteTime.Month)

                        # Create year and month directories with validation
                        $yearPath = Join-Path -Path $Data2SortPath -ChildPath $year
                        $monthPath = Join-Path -Path $yearPath -ChildPath $monthName

                try 
                    {
                        if (-not (Test-Path -Path $yearPath)) 
                            {
                                New-Item -ItemType Directory -Path $yearPath -Force | Out-Null
                                    Write-MoveLog "Created year folder at '$yearPath'." "INFO"
                            }
                        if (-not (Test-Path -Path $monthPath)) 
                            {
                                New-Item -ItemType Directory -Path $monthPath -Force | Out-Null
                                    Write-MoveLog "Created month folder at '$monthPath'." "INFO"
                            }
                    } 
                catch 
                    {
                                Write-MoveLog "Failed to create directory structure for '$file.FullName': $_" "ERROR"
                            continue
                    }

                        # Move file to the appropriate folder
                        $destination = Join-Path -Path $monthPath -ChildPath $file.Name

                        try 
                            {
                                if (-not (Test-Path -Path "$($destination)")) 
                                    {
                                        Move-Item -Path $file.FullName -Destination $destination -Verbose
                                            Write-MoveLog "File '$($file.FullName)' moved to '$duplicateName'." "INFO"
                                    }
                                if (Test-Path -Path "$($destination)\$($file.Name)") 
                                    {
                                        $duplicateName = Join-Path $duplicatesPath -ChildPath ("{0}_{1}{2}" -f $file.BaseName, $counter, $file.Extension)
                                        $counter++

                                        Move-Item -Path $file.FullName -Destination $duplicateName -Verbose
                                            Write-MoveLog "File '$($file.FullName)' moved to '$duplicateName'." "WARNING"
                                   }
                            } 
                        catch 
                            {
                                    Write-MoveLog "Failed to move file '$file.FullName': $_" "ERROR"
                                continue
                            }
            
                    }
            } 
        catch 
            {
                # Log and continue on error
                Write-MoveLog "Error processing file '$($file.FullName)': $_" "ERROR"
                continue
            }
    }

# Log script completion
Write-MoveLog "File organization complete. See details in the log file at '$LogFile'." "INFO"
Write-Host "Processing complete. Logs are saved to '$LogFile'."


