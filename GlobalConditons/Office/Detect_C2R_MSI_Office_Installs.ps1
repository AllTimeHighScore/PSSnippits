#####################################
#Detect and  Office Version C2R or MSi install
#Author Kevin Van Bogart
#This ended up not fitting the need, but kept it for future use as a time saver.
######################################

$RegPaths = @()
$OfficeVersions = @()
$InstalledVersion = $null

#Check OS Architecture
$RegPaths = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall')
if ([environment]::Is64BitOperatingSystem){
    $RegPaths += 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
}

#Matches only the base install - 0011 = 0{2}1{2} with no language code (0000 or 0{4})
[regex]$Regex = "^{90([1-9]{1})([0-9]{1})(\d{4})-0{2}1{2}-(0{4})-[0-1]0{3}-0{7}FF1CE}$"

ForEach($RegPath in $RegPaths){
    (Get-ChildItem -Path $RegPath).PSChildName | ForEach-Object {
        if ($OfficeGUID = ($Regex.Matches($_)).value){
            #Concatenate to avoid corrupting the variable
            $OfficeVersions += (Get-ItemProperty -Path "$RegPath\$OfficeGUID").DisplayVersion
        }
    }
}#End check for C2R

if ($OfficeVersions[0] -ne $null){
    $InstalledVersion = $OfficeVersions[0]
}

if ($InstalledVersion -eq $null){
    #Assume no MSI version is installed if installed version
    #Universal Click2Run Check
    $OfficeC2R = $null
    $C2RVersion = $null
    
    #Initialize this as a var just to shorten the commands about to be run
    #Will always populate in native achitecture reg path
    $Path = "HKLM:SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
    
    #Visio and Project populate in this location with a comma delimited list. 'REG_SZ'
    #One of those products could exist while the MSI version of Office is present.
    if ($C2RProductIDs = (Get-ItemProperty -Path $Path -ErrorAction SilentlyContinue).ProductReleaseIds){
        $OfficeC2R = $C2RProductIDs.split(',') | Where-Object {$_ -match "Office|^O365ProPlus"}
        
        #Grab the C2R Version. The versions should always match office, even if Visio and Project are installed. 
        if (($OfficeC2R -ne $null) -and ($C2RVersion = (Get-ItemProperty -Path $Path -ErrorAction SilentlyContinue).VersionToReport)){
            $InstalledVersion = $C2RVersion
        }
    }#End check for Office C2R Prod ID
}#End checking if version already populated

if ($InstalledVersion -ne $null){
    $InstalledVersion
}
else {
    '0.0.0.0'
}
