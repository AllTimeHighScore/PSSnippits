################
# Match Uninstall Product Code Display Name - Where method
# Kevin Van Bogart
# Suitable as a custom CM detection method in SCCM
# - Item(s) will only be in one hive or the other.
###############

$DisplayName = 'Configuration Manager Client'

if ($InitData.ENVx64OS){
    $RegHive ='HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
}
else {
    $RegHive ='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
}

$ProductLocated = Get-ChildItem -Path $RegHive -ErrorAction SilentlyContinue |
    Get-ItemProperty -Name DisplayName -ErrorAction SilentlyContinue |
        Where-Object {$_.DisplayName -match $DisplayName} |
            Select-object -ExpandProperty PSchildname

if ($ProductLocated){'Valid'}



################
# - Match Uninstall Product Code Display Name - Where method
# - Kevin Van Bogart
# - Suitable as a custom CM detection method in SCCM
# - Item(s) may be in either hive
###############

$DisplayName = 'Configuration Manager Client'

$RegHive = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall')

if ((Get-CimInstance -ClassName Win32_OperatingSystem).OSArchitecture -match '64'){
    $RegHive +='HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
}

$ProductLocated = foreach ($Hive in $RegHive){
    Get-ChildItem -Path $Hive -ErrorAction SilentlyContinue |
        Get-ItemProperty -Name DisplayName -ErrorAction SilentlyContinue |
            Where-Object {$_.DisplayName -match $DisplayName} |
                Select-object -ExpandProperty PSchildname
}

if ($ProductLocated){'Valid'}

################
# - Match Uninstall Product Code Display Name - Where method
# - Kevin Van Bogart
# - Suitable as a custom CM detection method in SCCM
# - Item(s) may be in either hive and have nicer output.
###############

$DisplayName = '*McAfee*'
$Prefilter = 'Designer'
$McAfeeProducts = $null

$RegHive = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall')

#Assuming no module, there's a few different ways you could go here. I chose this.
if ((Get-CimInstance -ClassName Win32_OperatingSystem).OSArchitecture -match '64'){
    $RegHive +='HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
}

#The -ne $null is used because some software with an empty displayname can cause false positives, such as Turbo Tax... The bastards.
$Products += foreach ($Hive in $RegHive){
    Get-ChildItem -Path $Hive -ErrorAction SilentlyContinue | ForEach-Object { 
            Get-ItemProperty -Path $($_.pspath) -ErrorAction SilentlyContinue | 
                Where-Object { ($_.DisplayName -match $DisplayName) -and ($_.DisplayName -ne $null) -and ($_.DisplayName -notmatch $PreFilter) } | 
                    ForEach-Object {

                        [pscustomobject]@{
                            KeyName = $(Split-Path -Path $_.PsPath -leaf)
                            KeyPath = ($_.PsPath).replace('Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE','HKLM:')
                            DisplayName = $_.DisplayName
                            DisplayVersion = $_.DisplayVersion
                            UninstallString = $_.UninstallString
                            WOW6432Node = if ($_.PsPath -match 'WOW6432Node'){$true}else{$false}
                        }
                    }
        }

}

$Products