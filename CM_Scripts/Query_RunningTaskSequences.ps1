
$Query = "SELECT * FROM CCM_TSExecutionRequest WHERE RunningState = 'NotifyExecution'"
#$Query = "SELECT * FROM CCM_ExecutionRequestEx WHERE State = 'Running'"
#$RunningShit = Get-ciminstance -Query $Query -Namespace root\CCM\SoftMgmtAgent

$MemberProgs = (Get-ciminstance -Query $Query -Namespace root\CCM\SoftMgmtAgent).TS_MemberProgramID

$PKGAdverts = (Get-ciminstance -Query $Query -Namespace root\CCM\SoftMgmtAgent).OptionalAdvertisements

