
$group = "ADGroupName"

((New-Object System.DirectoryServices.DirectorySearcher("(&(objectCategory=User)(samAccountName=$env:USERNAME))")).FindAll()).Properties.memberof | 
    ForEach-Object { ([adsi]"LDAP://$_").cn } | ? { ($_ -eq $group) } 

