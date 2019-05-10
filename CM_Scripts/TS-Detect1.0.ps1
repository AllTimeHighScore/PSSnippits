# if the registry key is present then we'll send the correct response back for a DT to trigger on.

<#This was just a test for a much more involved script#>

$TaskSequenceFlag = (get-itemproperty -path HKLM:\SOFTWARE\BSC\TaskSequences -Name Running -ErrorAction SilentlyContinue).Running
$TaskSequenceService = (get-service -Name smstsmgr).status

if (($TaskSequenceFlag -eq '1') -and ($TaskSequenceService -eq 'running')){
    1
}
else {
    0
}