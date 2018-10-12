
$TargetUser = 'DOMAIN\User'
if ((Get-LocalGroupMember -Group 'Administrators').Name -contains $TargetUser){
    'User or User group already Present'
}