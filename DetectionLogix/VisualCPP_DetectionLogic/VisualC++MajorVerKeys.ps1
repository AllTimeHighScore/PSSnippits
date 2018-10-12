#################################
# - Visual C++ Universal Detection Logic: Major Version Keys
# - Author: KVB 08-02-2018
# - Referrence: https://blogs.msdn.microsoft.com/astebner/2010/05/05/mailbag-how-to-detect-the-presence-of-the-visual-c-2010-redistributable-package/
#               https://msdn.microsoft.com/en-us/library/ms235299.aspx
#################################
[version]$CurrentVersion = '14.10.25017.0'
$Paths = $null
$VCInstalls = @()

if ([environment]::Is64BitOperatingSystem){
    $Paths += @(
        "HKLM:\SOFTWARE\Wow6432Node\Microsoft\VisualStudio\$($CurrentVersion.Major).0\VC\Runtimes\x64",
        "HKLM:\SOFTWARE\Wow6432Node\Microsoft\VisualStudio\$($CurrentVersion.Major).0\VC\Runtimes\x86"
    )
}
else {
    $Paths = "HKLM:\SOFTWARE\Microsoft\VisualStudio\$($CurrentVersion.Major).0\VC\Runtimes\x86"
}

$Paths | ForEach-Object {
    #Encompasing the object comparison because Get-itemproperty will send a blank objects through the pipeline
    $VCInstalls += (Get-ItemProperty -Path $_ -ErrorAction SilentlyContinue).Version | Where-Object {if ($_){[version]$_.replace('v','') -ge $CurrentVersion}}
}

#Compare the items with the paths
if ($Paths.count -eq $VCInstalls.count){
    'Compliant'
}