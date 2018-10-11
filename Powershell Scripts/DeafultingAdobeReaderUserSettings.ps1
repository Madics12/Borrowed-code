#American Fidelity Assurance
#XenApp 32-bit User Logon Script for workstation images

#Functions Definitions
#****************************************************************************************************************************************************************
#Return powershell scripts folder path
#****************************************************************************************************************************************************************
Function Get-ScriptDirectory 
{
 $Invocation = (Get-Variable MyInvocation -Scope 1).Value
 Split-Path $Invocation.MyCommand.Path 
}


#****************************************************************************************************************************************************************
#Declare Variables
#****************************************************************************************************************************************************************
$strAdobeReaderEula = "HKCU:\Software\Adobe\Acrobat Reader\DC\AdobeViewer"


#****************************************************************************************************************************************************************
#Supress Adobe Reader first-run settings
#****************************************************************************************************************************************************************


if (!(Test-Path -Path "HKCU:\Software\Adobe\Acrobat Reader")) {
  New-Item -Path "HKCU:\Software\Adobe" -Name "Acrobat Reader" -Force
  New-Item -Path "HKCU:\Software\Adobe\Acrobat Reader" -Name "DC" -Force
  New-Item -Path "HKCU:\Software\Adobe\Acrobat Reader\DC" -Name "AdobeViewer" -Force
}

New-ItemProperty -Path $strAdobeReaderEula -Name "EULA" -Value "1" -PropertyType "DWORD" -Force
