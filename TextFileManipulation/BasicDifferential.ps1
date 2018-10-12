##############
# Desc: Basic Diff
# Author: Can you really call me the author(KVB)? This is lame
# Note: Only use this as a quick reference. I can do more with this.
#############

$File1 = Get-Content -Path (Join-Path -Path $PSScriptRoot -ChildPath 'OldCode\Tableau_Desktop_10.5.2_PKG_All.ps1')
$File2 = Get-Content -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Current\Tableau_Desktop_10.5.2_PKG_All.ps1')

$Compare = Compare-Object -ReferenceObject $File1 -DifferenceObject $File2 -PassThru  