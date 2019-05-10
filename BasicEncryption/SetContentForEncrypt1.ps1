#YOURMOTHER

$file = "\\sNETWORKDRIVE\McAfee_ENS_10.5_PKG_All\Staging\McRemfee.txt"

read-host -assecurestring | convertfrom-securestring -Key (1..16) | out-file $file

$Secure2 = Get-Content $file | ConvertTo-SecureString -Key (1..16)
