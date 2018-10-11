#Author: Stephen Trapanese
#This script will prompt for each parameter and then terminate the process for that user
#This is helpful when you have a RDS user that needs a program process ended remotely

$remotecomputer = Read-Host "Enter Computer Name"
$process_name = Read-Host "Enter Process Name example:notepad.exe"
$user = Read-Host "Enter username to terminate the process for"


Get-WmiObject -Class Win32_Process -Filter "Name= '$process_name'" -ComputerName $remotecomputer | 
Where-Object { $_.GetOwner().User -eq $user } | 
Foreach-Object { $_.Terminate() }