
#([wmi]'ROOT\ccm\clientsdk:CCM_Application.Id="ScopeId_E5D6EF33-13ED-4746-B53F-F50165723E29/Application_c402fd6f-de13-454e-9977-9820e8d8f9ef",IsMachineTarget=TRUE,Revision="6"').AppDTs

#$Query = "Select * from CCM_Application where Name like 'Shockwave'"
$Query = "Select * from CCM_Application where Id like '%Application_63d1aadb-1e98-4410-9e78-eb9e1032ccbb'"#ens
$Query = "Select * from CCM_Application where Id like '%Application_a0c2668d-b584-44c3-a0b0-70e624f73729'"#drive encryption 7.2.4


$targetApp = Get-ciminstance -Query $Query -Namespace 'ROOT\ccm\clientsdk' 

$targetApp