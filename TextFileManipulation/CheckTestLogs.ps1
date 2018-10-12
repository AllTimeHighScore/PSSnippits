###################
# Scan log files for result codes
# Note: This can change wildly depending on what you're looking for
#     This is ultra basic, but that's all I typically want for this sort of task.
# Author: Kevin Van Bogart
###################

<#
***NOTE: Something is wrong here. It seems to miss certain exit codes.
#>


$AllExits = $null
$ErrCodes = $null

$TargetDir = Join-Path -Path $PSScriptRoot -ChildPath 'YourLogDir'

#Search the directory
$Logs = Get-ChildItem -Path $TargetDir -File -Recurse | ? {$_.FullName -notmatch "-install|-uninstall|msi|txt"} | % {Get-content -path $_.FullName}

#Get the full name of the logs
$AllExits = $Logs | ? {$_ -match "EXIT CODE:"}
$ErrCodes = $Logs | ? {($_ -match "EXIT CODE:") -AND ($_ -notmatch "EXIT CODE: 0|EXIT CODE: 3010")}
$Logs | ? {$_ -match "EXIT CODE:"}

#Spit out any error codes. The rest can be inspected manually.
$ErrCodes