#
$fld = "C:\Logs\Folder1"
$fld2 = "C:\Logs\Folder2"
$fil = '*.*'

$fsw = New-Object System.IO.FileSystemWatcher $fld, $fil -Property @{includeSubDirectories= $true;notifyFilter = [IO.notifyFilters] 'filename,lastwrite,lastaccess,directoryname'}
$fsw2 = New-Object System.IO.FileSystemWatcher $fld2, $fil -Property @{includeSubDirectories= $true;notifyFilter = [IO.notifyFilters] 'filename,lastwrite,lastaccess,directoryname'}
New-EventLog -LogName FolderMon -Source

#monitor folder 1 for object created
Register-ObjectEvent $fsw created -SourceIdentifier filecreated -Action {
$name = $event.sourceEventArgs.name
$changeType = $event.sourceEventArgs.ChangeType
$timeStamp = $event.timeGenerated

Write-Host "$name was $changeType at $timeStamp" -ForegroundColor Red
}

Register-ObjectEvent $fsw renamed -SourceIdentifier filecRenamed -Action {
$name = $event.sourceEventArgs.name
$changeType = $event.sourceEventArgs.ChangeType
$timeStamp = $event.timeGenerated

Write-Host "$name was $changeType at $timeStamp" -ForegroundColor yellow
}

#monitor folder 2 for object created
Register-ObjectEvent $fsw2 created -SourceIdentifier filecreated2 -Action {
$name2 = $event.sourceEventArgs.name
$changeType2 = $event.sourceEventArgs.ChangeType
$timeStamp2 = $event.timeGenerated

Write-Host "$name2 was $changeType2 at $timeStamp2" -ForegroundColor Red

}

#stops monitoring folders
pause
Unregister-Event filecreated
Unregister-Event filecreated2
Unregister-Event filerename


