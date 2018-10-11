Import-Module ActiveDirectory

$content = Get-Content "UserAccounts.txt"

foreach ($line in $content)
{
   #Write-Output $line
#   $useraccount = Get-ADUser -Properties Department -filter {SamAccountName -like $line} | Select-Object department
   $useraccount = Get-ADUser -filter {SamAccountName -like $line} | Select-Object Name

   if ($useraccount.Name) { 
    Write-Output $useraccount.Name
   }
   else { Write-Output " " }
   #Write-Output `n
  
}
