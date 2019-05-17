$Packages = get-wmiobject -query "SELECT * FROM CCM_SoftwareDistribution" -namespace "root\ccm\policy\machine\actualconfig"
$Apps = $null
#($Packages | Where-Object {$_.PKG_Name -match "KEVINTEST_TaskSequenceToolKit_1.0_PKG_All"}).PKG_Name
#($Packages | Where-Object {$_.OptionalAdvertisements -match "CM220263"}).OptionalAdvertisements
$Selected = $Packages | Where-Object {$_.ADV_AdvertisementID -match "CM220263"}
$MemberProgs = $Selected.TS_References
$Apps = ($MemberProgs | 
        Where-Object {($_ -match "Application_*")} -ErrorAction SilentlyContinue | 
        ForEach-Object {$_.replace('>','').replace('<','').replace('"','').Split('/')} | 
        Get-Unique).where({$_ -match "Application_"},'SkipUntil')
$Apps
#.where({$_ -match "Application_"},'SkipUntil')
#ADV_AdvertisementID = CM123456
#PKG_PackageID = CM789123

#OptionalAdvertisements   : {CM123456}
#ContentID                : CM789123


"ADV Run Notification = $($Selected.ADV_ADF_RunNotification)"
"PRG Run Notification = $($Selected.PRG_PRF_RunNotification)"
"Working dir = $($Selected.PRG_WorkingDirectory)"

(get-item -path $PSCOMMANDPATH).LastWriteTimeUtc

