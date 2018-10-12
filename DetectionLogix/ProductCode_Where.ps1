################
# Match Uninstall Product Code Display Name - Where method
# Kevin Van Bogart, AKA : Agent K, KVB, Veebster
# Suitable as a custom CM detection method in SCCM
# *** Warning *** not as acurate as a GUID search. Be very sure the name of the target app.
###############

$DisplayName = 'Configuration Manager Client'

if ($InitData.ENVx64OS){
    $RegHive ='HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
}
else {
    $RegHive ='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
}    

$ProductLocated = Get-ChildItem -Path $RegHive -ErrorAction SilentlyContinue | 
    Get-ItemProperty -Name 'DisplayName' -ErrorAction SilentlyContinue | 
        Where-Object {$_.DisplayName -match $DisplayName} | 
            Select-object -ExpandProperty PSchildname


if ($ProductLocated){'Valid'}

