#Universal Click2Run Check
$OfficeC2R = $null
$C2RVersion = $null

#Initialize this as a var just to shorten the commands about to be run
#Will always populate in native achitecture reg path
[string]$Path = "HKLM:SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
[string]$Bitness = $null

#Visio and Project populate in this location with a comma delimited list. 'REG_SZ'
#One of those products could exist while the MSI version of Office is present.
if ($C2RProductIDs = (Get-ItemProperty -Path $Path -ErrorAction SilentlyContinue).ProductReleaseIds){
    $OfficeC2R = $C2RProductIDs.split(',') | Where-Object {$_ -match "Office|^O365ProPlus"}
    $Type = "C2R"
    #Grab the C2R Version. The versions should always match office, even if Visio and Project are installed. 
    if ($OfficeC2R -and ([version]$C2RVersion = (Get-ItemProperty -Path $Path -ErrorAction SilentlyContinue).VersionToReport)){
        $Platform = (Get-ItemProperty -Path $Path -ErrorAction SilentlyContinue).Platform
        $Family = $C2RVersion.Major
    }
}

If ($Platform -eq 'x64'){
    $Bitness = '64'
}
elseif ($Platform -eq 'x86'){
    $Bitness = '32'
}

-Join($Type,$Bitness,$Family)