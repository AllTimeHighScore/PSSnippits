                
$PKG_Query = "SELECT * FROM CCM_SoftwareDistribution where ADV_AdvertisementID like '%CM220263%'"

$Packages = get-ciminstance -query $PKG_Query -namespace "root\ccm\policy\machine\actualconfig" -ErrorAction Stop

#Look for the actual packages
$TS_PKGs = $Packages | Where-Object {!($_.TS_Sequence)} -ErrorAction SilentlyContinue | Get-Unique

$TS_PKGs | ForEach-Object {
    

    [string]$CommandLine = $_.PRG_CommandLine
    #$CommandLine = '%winDir%\Sysnative\WindowsPowerShell\v1.0\powershell.exe -NoLogo -NonInteractive -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -file PauseTaskSequence.ps1'
    
    #$CommandLine.split(' ').where({($_ -in '.ps1','.exe','.msi','.msp','.vbs','.msu') -and ($_ -notin 'cmd','wscript','powershell','wusa')})
    if ($Commandline){
        $Inculde = "\.ps1|\.exe|\.msi|\.msp|\.vbs|\.msu|\.bat"
        $Exclude = "cmd|wscript|powershell|wusa|^-"
        $PkgFile = $CommandLine.split(' ').split('\').where({($_ -match $Inculde) -and ($_ -notmatch $Exclude)}).split('.')[0]
        $PkgFile
    }
}