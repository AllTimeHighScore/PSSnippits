


$file = "\\NETWORKDRIVE\ZScaler_ZScaler_1.2.4_PKG_All\Staging\UninstallPassword.txt"


$pass = Get-Content $file | ConvertTo-SecureString -Key (1..16)

$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($pass)
$UnsecurePassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)


