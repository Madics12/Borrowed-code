function Get-ScriptDirectory{
 if($hostinvocation -ne $null) {
 Split-Path $hostinvocation.MyCommand.path
 }
 else {
 (Get-Location -PSProvider FileSystem).ProviderPath
 }
}

$scriptdir = Get-ScriptDirectory
$offlineLog = "$scriptdir\OfflineComputers.txt"
$wksquerylog = "$scriptdir\WorkstationQueryLog.txt"

$reader = [System.IO.File]::OpenText("$scriptdir\Workstations.txt")

try {
    for(;;) {
        $workstation = $reader.ReadLine()
        if ($workstation -eq $null) { break }
        # process the line
        
        "Workstation queried: $workstation" | Add-Content $wksquerylog
	 Write-Host "Workstation queried: $workstation"

        if (Test-Connection -ComputerName $workstation -Quiet -Count 2 -ErrorAction SilentlyContinue) {
            Restart-Computer -ComputerName $workstation -Force
            Start-Sleep -s 30
        }
        else {
         #computer is offline
          "Workstation is Offline: $workstation" | Add-Content $offlineLog
        }


     }#end for

}#end try
finally {
    $reader.Close()
}