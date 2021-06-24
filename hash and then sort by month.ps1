<#
Update the path of Import and Duplicates. Import is where all files and subfolders are located that require sorting.

There are 2 sections to the script that can be executed independently or as a single script. The first compares file hashes and moves any duplicates to the duplicate folder. The second part sorts the files by date creating a year folder and a month subfolder.


Although tested this has not been rigorously tested and I take no responsibility for any data loss. So be careful. 
#>

$path = "D:\Images\import"
$duplicates = "D:\Images\duplicates"

New-Item $duplicates -ItemType Directory -Force

$files = Get-ChildItem -Path $path -Recurse -File

$fileHashes = @()

foreach($file in $files)
{
    $thisFileHash = Get-FileHash -Algorithm SHA256 -Path $file.FullName

    if ($fileHashes.Count -eq 0)
    {
        $fileHashes += @($thisFileHash.Hash)
    }
    else
    {
        $isDuplicate = $false
        foreach ($fileHash in $fileHashes)
        {
            if($fileHash -eq $thisFileHash.Hash)
            {
                $isDuplicate = $true
                break
            }
        }
        if($isDuplicate)
        {
           $nextName = Join-Path -Path $duplicates -ChildPath $_.name 

        while(Test-Path -Path $nextName)
        {
       $nextName = Join-Path $duplicates ($file.BaseName + "_$num" + $file.Extension)    
       $num+=1   
    }

            Move-Item -Path $file.FullName -Destination "$nextname" -Verbose
        }
        else
        {
            $fileHashes += @($thisFileHash.Hash)
        }
    }
}


$file = Get-ChildItem $path -Recurse -File | Select-Object *
#files that are moved to folders based on year and month
foreach ($fi in $file)
{
    $month = $fi.lastwritetime.month
    $year = $fi.LastWriteTime.Year
    $fullname = $fi.FullName

    $monthName = (Get-Culture).DateTimeFormat.GetMonthName($month)
   # $yrMonth = "$year"+"_"+"$monthName"

    New-Item -ItemType Directory -Path $path\ -name $year -ErrorAction SilentlyContinue
    New-Item -ItemType Directory -Path $path\$year -name $monthName -ErrorAction SilentlyContinue
    $destination = "$path\$year\$monthName"
    Move-Item $fullname $destination
}