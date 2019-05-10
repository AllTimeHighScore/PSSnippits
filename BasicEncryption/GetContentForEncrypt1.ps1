$PWFile = "\\stpnas05\dsm\Dev\Packages\McAfee_ENS_10.5_PKG_All\Staging\McRemfee.txt"

$SecurePW = Get-content $PWFile | ConvertTo-SecureString -Key (1..16)

$UnsecurePassword = (New-Object PSCredential "user",$SecurePW).GetNetworkCredential().Password



#[Byte[]] $key = (1..16)
##$Unsecured = 
#(Get-Content $PWFile | ConvertTo-SecureString -Key $key)


