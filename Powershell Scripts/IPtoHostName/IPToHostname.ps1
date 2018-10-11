$content = Get-Content "IPList.txt"

foreach ($ipadd in $content)
{ 
    
   $currentEAP = $ErrorActionPreference 
   $ErrorActionPreference = "silentlycontinue" 

   $Result = [System.Net.Dns]::gethostentry($ipadd)
   $ErrorActionPreference = $currentEAP
   
   If ($Result) { 
      Write-Host $Result.HostName
      $Result.HostName | Add-Content "Out.txt"
   } 
   Else 
   { 
      Write-Host "$IP – No HostNameFound" 
      "$IP – No HostNameFound" | Add-Content "Out.txt"
   } 

}
