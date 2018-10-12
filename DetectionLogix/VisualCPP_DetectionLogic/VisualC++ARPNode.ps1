#############################################
# - Find Instances of Visual C++ redistributable installed 
# - I really don't like looking at this location for this application 
#  Not all components install in each hive, so it might be best to only 
#  use it for redistributables. That's largely because you can't easily 
#  tally the components installed. Other methods are better.
#  Some older versions are a bit wacky with how they show up in the list.
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
