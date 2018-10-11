Import-Module AppvClient

function Get-ScriptDirectory{
 if($hostinvocation -ne $null) {
 Split-Path $hostinvocation.MyCommand.path
 }
 else {
 (Get-Location -PSProvider FileSystem).ProviderPath
 }
}

$scriptdir = Get-ScriptDirectory
$User = $env:Username


$SDrive = Get-PSDrive "S"

if (($SDrive -eq $null) -or ($SDrive.DisplayRoot -ne "\\ffffg-pp-file01\shares")) { 
	if ($SDrive.DisplayRoot -ne "\\ffffg-pp-file01\shares") {
	  Remove-PSDrive -Name 'S' -Force
	  Start-Sleep -s 1
	}
	
	New-PSDrive -Name "S" -PSProvider FileSystem -Root "\\ffffg-pp-file01\shares" -Persist
}

if (Test-Path "S:\COBRA Solutions\COBRA Administration Manager\Cobra.exe") {

  try {
    $pkgid = get-appvclientpackage -Name "*CobraSolutions*" | select-object -ExpandProperty PackageId
    $vid = Get-AppvClientPackage -PackageId $pkgid | Select-Object -ExpandProperty VersionId
    
    $appvarg = "/appvve:$pkgid" + "_" + "$vid"
    Start-Process "S:\COBRA Solutions\COBRA Administration Manager\Cobra.exe" -ArgumentList "$appvarg" -WorkingDirectory "S:\COBRA Solutions\COBRA Administration Manager" 

  }

  Catch { }
  
}
else {
  $wshell = New-Object -ComObject Wscript.Shell -ErrorAction Stop
  $wshell.Popup("Unable to get to the Cobra executable.",10,"Status",1)
}