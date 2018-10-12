####################################
# - Returns the version of Office
# -    regardless of how it's installed
# -    Example: MSI or Click-To-Run (C2R)
# - Author: Kevin Van Bogart
# - Created 07-06-2018
####################################

#Add reg paths
$RegPaths = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall')
switch ($true){
    ([environment]::Is64BitOperatingSystem) {$RegPaths += 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'}
    (Test-Path -Path 'HKLM:SOFTWARE\Microsoft\Office\ClickToRun') {$RegPaths += 'HKLM:SOFTWARE\Microsoft\Office\ClickToRun'}
}

#Send each registry path through a loop for inspection
ForEach($RegPath in $RegPaths){

    #Inspect the child keys
    Get-ChildItem -Path $RegPath -PipelineVariable 'Key' -ErrorAction SilentlyContinue | ForEach-Object {
    
        #Check The Keys
        #Regex match for 64-bit Office ProPlus base installs
        if ($Key.PSChildName -match '^{90([1-9]{1})([0-9]{1})(\d{4})-0{2}1{2}-0{4}-1{1}0{3}-[0-1]0{6}FF1CE}$'){
            $Version = (Get-ItemProperty -Path "$RegPath\$($Key.PSChildName)" -ErrorAction SilentlyContinue).DisplayVersion
        }
        #Regex match for 32-bit Office ProPlus base installs
        elseif ($Key.PSChildName -match '^{90([1-9]{1})([0-9]{1})(\d{4})-0{2}1{2}-(0{4}-){2}[0-1]0{6}FF1CE}$'){
            $Version = (Get-ItemProperty -Path "$RegPath\$($Key.PSChildName)" -ErrorAction SilentlyContinue).DisplayVersion
        }
        #Look for the C2R sub keys
        #Add an additional sanity check in case some vendor did something insane like use an uninstall key named 'Configuration'
        elseif (($Key.PSChildName -match '^Configuration$') -and ($Key.PSPath -match 'ClickToRun')){
            $Version = (Get-ItemProperty -Path "$RegPath\$($Key.PSChildName)" -ErrorAction SilentlyContinue).VersionToReport
        }#end else configuration found
    }#End ChildItem loop
}#Ens regpath loop

$Version