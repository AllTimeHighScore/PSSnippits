$Packages = get-wmiobject -query "SELECT * FROM CCM_SoftwareDistribution" -namespace "root\ccm\policy\machine\actualconfig"
$Apps = $null
#($Packages | Where-Object {$_.PKG_Name -match "EXP_PKG_BSC_TaskSequenceToolKit_1.0_PKG_All"}).PKG_Name
#($Packages | Where-Object {$_.OptionalAdvertisements -match "CM220263"}).OptionalAdvertisements
$Selected = $Packages | Where-Object {$_.ADV_AdvertisementID -match "CM220263"}
$MemberProgs = $Selected.TS_References
$Apps = ($MemberProgs | 
        Where-Object {($_ -match "Application_*")} -ErrorAction SilentlyContinue | 
        ForEach-Object {$_.replace('>','').replace('<','').replace('"','').Split('/')} | 
        Get-Unique).where({$_ -match "Application_"},'SkipUntil')
$Apps
#.where({$_ -match "Application_"},'SkipUntil')
#ADV_AdvertisementID = CM220263
#PKG_PackageID = CM200B63

#OptionalAdvertisements   : {CM220263}
#ContentID                : CM200ABD


"ADV Run Notification = $($Selected.ADV_ADF_RunNotification)"
"PRG Run Notification = $($Selected.PRG_PRF_RunNotification)"
"Working dir = $($Selected.PRG_WorkingDirectory)"

(get-item -path $PSCOMMANDPATH).LastWriteTimeUtc

