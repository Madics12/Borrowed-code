import-module ActiveDirectory

$groups = Get-ADGroup -filter {name -like "VDI-*"}

foreach ($group in $groups) { Write-Host $group.DistinguishedName " : " (Get-ADGroupMember $group.DistinguishedName).count }