
[xml]$xml = Get-Content C:\Powershell\XML_Test\Public.xml

foreach ($entity in $xml.Public.Person){
$name = $entity.Name
$age = $entity.Age

    if($age -gt 60){
        echo "Hello $name You are qualified for a senior citizen quota"
    } else {
        echo "Hello $name You are NOT qualified for senior citizen quota"
        }
    }
    