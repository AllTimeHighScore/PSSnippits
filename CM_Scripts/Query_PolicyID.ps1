
#From this the advert and content ID's can be found.
#AdvertID = CM220000
#ContentID = CM200B34

$Query = "Select * from ccm_policy where policyid like '%e217b742-5cc3-4559-8aa0-8b9e3f8d510f'"
#$RunningShit = Get-ciminstance -Query $Query -Namespace root\CCM\SoftMgmtAgent

$testquery = Get-ciminstance -Query $Query -Namespace 'Root\ccm\Policy\machine\RequestedConfig' 


#$MyInvocation

$PSBOUNDPARAMETERS