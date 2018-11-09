#########################
# - Pac File Inspection
# - Created by Kevin Van Bogart
# - Created because..Well this is easier.
#########################

#Entries to check for in pac file
$InspectionElements = @('192.168.0.0')

#Proxy settings root
$ProxyMgr = 'HKLM:\SYSTEM\ControlSet001\Services\iphlpsvc\Parameters\ProxyMgr'

#Get the path that should have the proxy location so the uri can be used
if ($FullPSPath = (Get-ChildItem -Path $ProxyMgr -recurse -ErrorAction SilentlyContinue | Where-Object {$_.Property -match 'AutoConfigUrl'}).pspath){
    "Pac file found here --> $FullPSPath"


    #Try to grab the file so it can be inspected 
    try {
        Invoke-WebRequest -uri (Get-ItemProperty -Path $FullPSPath -ErrorAction SilentlyContinue).'AutoConfigUrl' -OutFile 'C:\temp\PacCopyAttempt.txt' -ErrorAction Stop
        "Pac File copied"
    }
    catch {
        "Encounted issue copying Pac File."
    }
    
    try {
        #Get content
        $PacContent = Get-Content -Path $MovedPac
    }
    catch {
        "Could not get content from: $MovedPac"  
    }
    
    #If there is any data to inspect, check for items
    if ($PacContent){
        $InspectionElements | ForEach-Object {
            if ($MatchedLine = $PacContent -match $_){
                "#########Item Found on following line########"
                $MatchedLine
            }
        }
    }

}
else {
    "No Pack file located."
}