#$DeploymentID = 'cm220263'
$DeploymentID = 'cm220219'



#$SchedMessageQuery = "Select * from CCM_Scheduler_ScheduledMessage where ScheduledMessageID like '" + $DeploymentID + "%'"
$SchedMessageQuery = "Select * from CCM_TaskSequence"

$TS_Objects = Get-WmiObject -Query $SchedMessageQuery -Namespace root\ccm\policy\machine\actualconfig

#$Query = "SELECT * FROM CCM_ExecutionRequest WHERE State = 'Running'"
#
#$RunningShit = Get-WmiObject -Query $SchedMessageQuery -Namespace root\CCM\SoftMgmtAgent
#
