
$SchedMessageQuery = "Select * from CCM_Scheduler_ScheduledMessage"

$yourmums = Get-WmiObject -Query $SchedMessageQuery -Namespace root\ccm\policy\machine\actualconfig -ErrorAction Stop 

