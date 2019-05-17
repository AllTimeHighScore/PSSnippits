####################################
# - Returns the major version of Office,
# - the application bitness (32-bit\64-bit),
# - and the type of the install in regard
# - to the licensing model regardless
# - of how it's installed
# - Created on 06-24-2018
# - Create By: vanbogk
####################################

#Add reg paths
$RegPaths = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall')
switch ($true){
    ([environment]::Is64BitOperatingSystem) {$RegPaths += 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'}
    (Test-Path -Path 'HKLM:SOFTWARE\Microsoft\Office\ClickToRun'){$RegPaths += 'HKLM:SOFTWARE\Microsoft\Office\ClickToRun'}
}

#Send each registry path through a loop for inspection
ForEach($RegPath in $RegPaths){

    #Inspect the child keys
    Get-ChildItem -Path $RegPath -PipelineVariable 'Key' -ErrorAction SilentlyContinue | ForEach-Object {
    
        #Check The Keys
        #Regex match for 64-bit Office ProPlus base installs
        if ($Key.PSChildName -match '^{90([1-9]{1})([0-9]{1})(\d{4})-0{2}1{2}-0{4}-1{1}0{3}-[0-1]0{6}FF1CE}$'){
            $Family = ([version](Get-ItemProperty -Path "$RegPath\$($Key.PSChildName)" -ErrorAction SilentlyContinue).DisplayVersion).major
            $Bitness = 64
            $Type = 'VL' #VL could be the default
        }
        #Regex match for 32-bit Office ProPlus base installs
        elseif ($Key.PSChildName -match '^{90([1-9]{1})([0-9]{1})(\d{4})-0{2}1{2}-(0{4}-){2}[0-1]0{6}FF1CE}$'){
            $Family = ([version](Get-ItemProperty -Path "$RegPath\$($Key.PSChildName)" -ErrorAction SilentlyContinue).DisplayVersion).major
            $Bitness = 32
            $Type = 'VL' #VL could be the default
        }
        #Look for the C2R sub keys
        #Add an additional sanity check in case some vendor did something insane like use an uninstall key named 'Configuration'
        elseif (($Key.PSChildName -match '^Configuration$') -and ($Key.PSPath -match 'ClickToRun')){

            #Get bit(ness)
            $Bitness = (Get-ItemProperty -Path "$RegPath\$($Key.PSChildName)" -ErrorAction SilentlyContinue).Platform.Replace('x','').Replace('86','32')
            #Grab the major\family version
            $Family = ([version](Get-ItemProperty -Path "$RegPath\$($Key.PSChildName)" -ErrorAction SilentlyContinue).VersionToReport).major
            #Grab audience data
            $AudienceData = (Get-ItemProperty -Path "$RegPath\$($Key.PSChildName)" -ErrorAction SilentlyContinue).AudienceData.Split('::')[-1]

            #If multiple items are installed, such as ProplusRetail(Office) or VisioRetail, split them by comma and examine each object
            (Get-ItemProperty -Path "$RegPath\$($Key.PSChildName)" -ErrorAction SilentlyContinue).ProductReleaseIds.Split(',') | ForEach-Object {
                
                #Check for C2R Enterprise User licensing such as O365 Retail
                if ($_ -match '\w{0,1}[0-9]{0,3}ProPlus[0-9]{0,4}Retail$'){
                    $Type = 'OL'
                }
                #Check for C2R Volume licensing such as Office 2019 - This could probably be dispensed with... as VL could be the default
                elseif (($_ -match '\w{0,1}[0-9]{0,3}ProPlus[0-9]{0,4}Volume$') -or ($AudienceData -match 'LTS[B-C]{1}')){
                    $Type = 'VL'
                }

            }#ForEach
        }#end else configuration found
    }#End ChildItem loop
}#Ens regpath loop

-Join($Type,$Bitness,$Family)