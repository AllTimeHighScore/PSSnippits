################################
# Desc: To isolate selected items in CSV files exported from SQL queries
# and make them more human readable
# Author: Kevin Van Bogart 
# Note: Simple, but this is just a basic example that can be built upon.
#################################

$Csv =  Join-Path -Path $PSScriptRoot -ChildPath 'ActiveApplication_Examples.csv'   
$Items = Import-Csv -Path $Csv
$AppObjects = $Items.AssignmentName

<#This produces duplicates and is old code of mine (not the most elegant) Might not be a bad idea for exceedingly large csv files
as foreach is a lighter weight option#>
##Foreach ($App in ($AppObjects | Where-Object {$_ -notmatch "EXP|TSApp"})){
##
##    #Where-Object {$App -notmatch "EXP|TSApp"} 
##    $CmAppName = ($App -split "_PKG_|_SRV_|_WKS_|_All_").replace("App_","").Replace("_"," ")
##    $CmAppName[0] | Out-File -FilePath (Join-Path -Path $PSScriptRoot -ChildPath 'FriendlyAppNames.txt') -Append
##}

#An attempt to strip duplicates, the Regex is still rudimentary here
$Items.AssignmentName | 
    Where-Object {$_.AssignmentName -notmatch "EXP|TSApp"} | 
        ForEach-Object {(($_ -split "_(PKG|SRV|WKS|All)_").replace("App_","").Replace("_"," ")) | Select-Object -First 1} | 
            Out-File -FilePath (Join-Path -Path $PSScriptRoot -ChildPath 'FriendlyAppNames.txt') -Append
