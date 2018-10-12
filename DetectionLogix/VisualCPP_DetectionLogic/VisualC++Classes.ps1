############################
# - Detect Components for newer versions of VC++
# - Does not capture anything below VC++2012
# - Author VanBogk 8-3-2018
# - It's not clear to me that this is much better than 
#  looking in the uninstall product code section. It appears
#  this might be the way this middleware is moving though.
############################

[version]$CurrentVersion = '14.10.25017.0'
[version]$CurrentShortVersion = "$($CurrentVersion.Major).$($CurrentVersion.Minor).$($CurrentVersion.Build)"
$RegPath = "HKLM:\SOFTWARE\Classes\Installer\Dependencies"
$Arches = @('x86')
$Year = '2017'
switch ($true){([environment]::Is64BitOperatingSystem) {$Arches += 'amd64'}}
$ExpectedComponents = @('Additional','Minimum','Redistributable')
$InstalledComponents = @()
$GUID = '{[A-Z0-9]{8}-([A-Z0-9]{4}-){3}[A-Z0-9]{12}}$'

#Step through the componensts
foreach ($Component in $ExpectedComponents){
    
    #Check for bitness
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
                $DisplayNameMatch = "(?<Product>C\+\+\s?)$Year\s(?<Type>Redistributable?)\s(?<Arch>(\(|)$($Arch.Replace('amd','x'))(\)|)?)\s-\s(?<version>\d+\.\d+\.\d+(\.\d+)??)"

                $GUIDKeys.foreach({
                    #Grab the actual key in a separate variable for greater readability
                    $KeyInspection = (Get-ItemProperty -Path $_.pspath -ErrorAction SilentlyContinue).DisplayName
                    #Return the grouped objects
                    $CPlusVer = ([regex]::matches($KeyInspection,$DisplayNameMatch)).Groups.Where({$_.Name -eq 'Version'}).value
                    if ([version]$CPlusVer -ge $CurrentVersion){
                        $InstalledComponents += $CPlusVer
                    }
                })#$GUIDKeys.foreach
            }#End Elseif $GuidKeys

        }#if ($Component -eq 'Redistributable'){..}
        #Grab the other components
        else {
            $RuntimeVer = $null
            #Again, for readability
            $RunTimeRegex = "Microsoft.VS.VC_Runtime$Component(VSU|)_(x86|amd64),v$($CurrentVersion.Major)"
            $RunTimeKey = (Get-ChildItem -Path $RegPath -ErrorAction SilentlyContinue) | Where-Object {$_.Name -match $RunTimeRegex}

             #Because it appears the version in the other components appear to be a different version, Redefine currentversion if necessary.
             $RunTimeKey.PSPath | ForEach-Object {
                if ($RuntimeVer = ([version](Get-ItemProperty -Path $_ -ErrorAction SilentlyContinue).Version).Revision -eq '-1'){$CurrentVersion = $CurrentShortVersion}
            }

            if ($RunTimeKey.PSPath | Where-Object {($RuntimeVer = [version](Get-ItemProperty -Path $_ -ErrorAction SilentlyContinue).Version) -ge $CurrentVersion}){
                $InstalledComponents += $RuntimeVer
            }
        }#Else
    }#End Foreach arch
}#End Foreach Component

#Counts
if (($ExpectedComponents.count * $Arches.Count) -eq $InstalledComponents.count){
    'Compliant'
}
