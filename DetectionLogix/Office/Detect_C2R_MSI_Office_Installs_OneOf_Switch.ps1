#Initialize vars, bitness needs to be a [string] because of replace method usage
[string]$Type = $null
[string]$Bitness = $null
[string]$Family = $null

#Add reg paths
$RegPaths = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall')
switch ($true){
    ([environment]::Is64BitOperatingSystem) {$RegPaths += 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'}
    (Test-Path -Path 'HKLM:SOFTWARE\Microsoft\Office\ClickToRun') {$RegPaths += 'HKLM:SOFTWARE\Microsoft\Office\ClickToRun'}
}

#Send each registry path through a loop for inspection
ForEach($RegPath in $RegPaths){

    #Inspect the child keys
    Get-ChildItem -Path $RegPath -PipelineVariable 'Key' | ForEach-Object {
 
        #Check The Keys
        Switch -Regex ($Key.PSChildName){
            
            #Regex match for 64-bit Office ProPlus base installs
            '^{90([1-9]{1})([0-9]{1})(\d{4})-0{2}1{2}-0{4}-1{1}0{3}-[0-1]0{6}FF1CE}$' {
                $Family = ([version](Get-ItemProperty -Path "$RegPath\$($Key.PSChildName)").DisplayVersion).major
                $Bitness = 64
                $Type = 'MSI'
            } #Close match for 64-bit installs MSI installs of Office
            
            #Regex match for 32-bit Office ProPlus base installs
            '^{90([1-9]{1})([0-9]{1})(\d{4})-0{2}1{2}-(0{4}-){2}[0-1]0{6}FF1CE}$' {
                $Family = ([version](Get-ItemProperty -Path "$RegPath\$($Key.PSChildName)").DisplayVersion).major
                $Bitness = 32
                $Type = 'MSI'
            } #Close match for 32-bit installs MSI installs of Office  

            #Look for the C2R sub keys 
            '^Configuration$' {
                
                #Add an additional sanity check in case some vendor did something nsane like use an uninstall product code names 'Configuration'
                if ((Get-ItemProperty -Path "$RegPath\$($Key.PSChildName)").ProductReleaseIds -match 'O365ProPlus|Office' ){
                    $Family = ([version](Get-ItemProperty -Path "$RegPath\$($Key.PSChildName)").VersionToReport).major
                    $Bitness = (Get-ItemProperty -Path "$RegPath\$($Key.PSChildName)").Platform
                    $Type = 'C2R'
                }#End last sanity check
            }#Close Match for C2R Office Installs 
        }#Ens switch
    }#End ChildItem loop
}#Ens regpath loop



-Join($Type,$Bitness.Replace('x','').replace('86','32'),$Family)