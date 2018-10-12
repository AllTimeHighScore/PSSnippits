#############################
#  Visio Custom Requirement
#
#  Detects any known Visio Install (that I care about :))
#  
#  Returns 1 if installed; 0 if not found
#
#  KVB 7-11-2018
#############################

#Initialize main variable
$VisioDetected = $false

$RegPaths = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall')
switch ($true){
    ([environment]::Is64BitOperatingSystem){$RegPaths += 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'}
    (Test-Path -Path 'HKLM:SOFTWARE\Microsoft\Office\ClickToRun\Configuration'){$RegPaths += 'HKLM:SOFTWARE\Microsoft\Office\ClickToRun\Configuration'}
}

$RegPaths | ForEach-Object {

    if ((Get-ItemProperty -Path "$_\Office14.VISIO" -ErrorAction SilentlyContinue).DisplayName -Match '^Microsoft Visio (Professional|Standard) 2010$'){
        $VisioDetected = $true
    }
    elseif ($_ -match 'ClickToRun'){
 
        #If multiple items are installed, such as ProplusRetail(Office) or VisioRetail, split them by comma and examine each object
        (Get-ItemProperty -Path $_ -ErrorAction SilentlyContinue).ProductReleaseIds.Split(',') | ForEach-Object {
            #Check for C2R Enterprise User licensing such as VisioProRetail
            if ($_ -match '^Visio[0-9]{0,4}(Pro|Std)[0-9]{0,4}Retail$'){
                $VisioDetected = $true
            }
        }#ForEach
    }
}

if ($VisioDetected -eq $true){
    1
}
else {
    0
}