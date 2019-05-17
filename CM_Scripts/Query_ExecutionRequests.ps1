


$Query = "SELECT * FROM CCM_ExecutionRequestEx"
#$RunningShit = Get-ciminstance -Query $Query -Namespace root\CCM\SoftMgmtAgent

Get-ciminstance -Query $Query -Namespace root\CCM\SoftMgmtAgent

