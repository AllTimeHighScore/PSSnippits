#These are built with the observation that there's not much actual need to focus on the year portion of the VC++ Middleware.

#####################
# - Option 1 Runtimes Registry Location
#####################

#################################
# - Visual C++ Universal Detection Logic Options
# - Author: KVB 08-02-2018
# - Choose 1 or multiple
#################################
[version]$CurrentVersion = '14.10.25017.0'
$Paths = $null
$VCInstalls = @()

if ([environment]::Is64BitOperatingSystem){
    $Paths += @(
        "HKLM:\SOFTWARE\Wow6432Node\Microsoft\VisualStudiog\$($CurrentVersion.Major).0\VC\Runtimes\x64",
        "HKLM:\SOFTWARE\Wow6432Node\Microsoft\VisualStudiog\$($CurrentVersion.Major).0\VC\Runtimes\x86"
    )
}
else {
    $Paths = "HKLM:\SOFTWARE\Microsoft\VisualStudio\$MajorVersion\VC\Runtimes\x86"
}

$Paths | ForEach-Object {
    #Encompasing the object comparison because Get-itemproperty will send a blank objects through the pipeline
    $VCInstalls += (Get-ItemProperty -Path $_ -ErrorAction SilentlyContinue).Version | Where-Object {if ($_){[version]$_.replace('v','') -ge $CurrentVersion}}
}

#Compare the items with the paths
if ($Paths.count -eq $VCInstalls.count){
    'Compliant'
}

#####################
# - Option 2: Installer Dependencies in Classes Location
# - This method is important in case we care which specific elements are present
#####################


############################
# - Detect Components for newer versions of VC++
# - Does not capture anything below VC++2012
# - Author VanBogk 8-3-2018
# - It's not clear to me that this is much better than 
#  looking in the uninstall product code section. It appears
#  this might be the way this middleware is moving though.
############################

[version]$CurrentVersion = '14.10.25017.0'
$RegPath = "HKLM:\SOFTWARE\Classes\Installer\Dependencies"
$Arches = @('x86')
switch ($true){([environment]::Is64BitOperatingSystem) {$Arches += 'amd64'}}
$ExpectedComponents = @('Additional','Minimum','Redistributable')
$InstalledComponents = @()
$GUID = '{[A-Z0-9]{8}-([A-Z0-9]{4}-){3}[A-Z0-9]{12}}$'

foreach ($Component in $ExpectedComponents){
    foreach ($Arch in $Arches){

        #2017 seems to be different here
        if ($Component -eq 'Redistributable'){
            $Redistributable = ",,$Arch,$($CurrentVersion.Major).0,bundle"
            #For Versions 2017+
            if ($BundleVersion = [version](Get-ItemProperty -Path "$RegPath\$Redistributable" -ErrorAction SilentlyContinue).Version | Where-Object {$_ -ge $CurrentVersion}){
                #only Adding this here to do sometihng in the if statement
                $InstalledComponents += $BundleVersion
            }
            #For 2012 and 2013 vc++ versions the value is stored as such
            elseif ($GUIDKeys = Get-ChildItem -Path $RegPath -ErrorAction SilentlyContinue | Where-Object {$_.pspath -match $GUID}){
                #Variablized for readability
                $DisplayNameMatch = "(?<Product>C\+\+\s?)[0-9]{4}\s(?<Type>Redistributable?)\s(?<Arch>(\(|)$($Arch.Replace('amd','x'))(\)|)?)\s-\s(?<version>\d+\.\d+\.\d+(\.\d+)??)"
                
                $GUIDKeys.foreach({
                    #Grab the actual key in a separate variable for greater readability
                    $KeyInspection = (Get-ItemProperty -Path $_.pspath -ErrorAction SilentlyContinue).DisplayName
                    #Return the grouped objects
                    $CPlusVer = ([regex]::matches($KeyInspection,$DisplayNameMatch)).Groups.Where({$_.Name -eq 'Version'}).value
                    if ((([version]$CPlusVer).Major -eq $CurrentVersion.Major) -and (([version]$CPlusVer -ge $CurrentVersion))){
                        $InstalledComponents += $CPlusVer
                    }
                })
            }
        }
        #Grab the other components
        else {
            $RuntimeVer = $null
            #Again, for readability
            $RunTimeRegex = "Microsoft.VS.VC_Runtime$Component(VSU|)_(x86|amd64),v$($CurrentVersion.Major)"
            $RunTimeKey = (Get-ChildItem -Path $RegPath -ErrorAction SilentlyContinue) | Where-Object {$_.Name -match $RunTimeRegex}
            if ($RunTimeKey.PSPath | Where-Object {($RuntimeVer = [version](Get-ItemProperty -Path $_ -ErrorAction SilentlyContinue).Version) -ge $CurrentVersion}){
                $InstalledComponents += $RuntimeVer
            }
        }
    }#End Foreach arch
}#End Foreach Component

#Counts
if (($ExpectedComponents.count * $Arches.Count) -eq $InstalledComponents.count){
    'Compliant'
}

#####################
# - Option 3: Uninstall Product Codes
#####################

#############################################
# - Find Instances of Visual C++ redistributable installed 
# - I really don't like looking at this location for this application 
#  Not all components install in each hive, so it might be best to only 
#  use it for redistributables. That's largely because you can't easily 
#  tally the components installed. Other methods are better.
#  Some older versions are a bit whacky with how they show up in the list.
# - KVB 8-3-2018
#############################################
[version]$CurrentVersion = '7.10.25017.0'
#[version]$CurrentVersion = '14.10.25017.0'

$Year = '2005'
$Arches = @('x86')
$32bitRegPath = $null
$RegPaths = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall')
switch ($true){([environment]::Is64BitOperatingSystem) {$Arches += 'x64';$RegPaths += 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'}}
$InstalledComponents = @()

foreach ($Path in $RegPaths){
    foreach ($Arch in $Arches){
        $DisplayNameMatch = "(C\+\+\s)$Year\s\s?(\(?$($Arch)\)?\s)?(Redistributable)\s?(\(?$($Arch)\)?)?"
        #Go one deeper to check the version on something as old as VC++ 2005
        Get-ItemProperty -Path "$Path\*" -ErrorAction SilentlyContinue | 
            Where-Object {($_.DisplayName -match $DisplayNameMatch) -and ([version]$_.DisplayVersion -ge $CurrentVersion)} |
            ForEach-Object {
                $InstalledComponents += $_.DisplayVersion
            }
    }#Arches
}#RegPaths

$InstalledComponents

###################################
# - Option 4: File detection
###################################

####################################
# - Locate the files the redistributable puts in place
# - Haven't found the files for the minimum or additional runtime, but haven't looked very hard either
# - KVB 08-03-2018
####################################
[version]$CurrentVersion = '14.10.25017.0'
$FilePaths = @("$env:windir\System32") 
switch ($true){([environment]::Is64BitOperatingSystem) {$FilePaths += "$env:windir\SysWOW64"}}
$Finds = @()
#Redistributables = mfc$($CurrentVersion.Major)0.dll

foreach ($Path in $FilePaths){
    $Finds += [version](Get-ItemProperty -Path "$Path\mfc$($CurrentVersion.Major)0.dll" -ErrorAction SilentlyContinue).VersionInfo.ProductVersion | 
        Where-Object {$_ -ge $CurrentVersion} 
}

if ($Finds.Count -eq $FilePaths.Count){"Compliant"}





