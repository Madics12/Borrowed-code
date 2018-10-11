$content = Get-Content "WorkstationList.txt"


foreach ($line in $content)
{
  get-adcomputer $line -Properties * | select name,description,operatingsystem
  
}