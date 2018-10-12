
####################################
# - Locate the files the redistributabel puts in place
# - Haven't found the files for the minimum or additional runtime, but haven't looked very hard either
# - KVB 08-03-2018
####################################
[version]$CurrentVersion = '14.10.25017.0'
$FilePaths = @("$env:windir\System32") 
switch ($true){([environment]::Is64BitOperatingSystem){$FilePaths += "$env:windir\SysWOW64"}}
$Finds = @()
#Redistributables = mfc$($CurrentVersion.Major)0.dll

foreach ($Path in $FilePaths){
    $Finds += [version](Get-ItemProperty -Path "$Path\mfc$($CurrentVersion.Major)0.dll" -ErrorAction SilentlyContinue).VersionInfo.ProductVersion | 
        Where-Object {$_ -ge $CurrentVersion} 
}

if ($Finds.Count -eq $FilePaths.Count){"Compliant"}

