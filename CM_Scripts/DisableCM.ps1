#just keeping these simple commands for McAfee and Java testing
#anything else that a force deployment of a previous version would interfere with would also benefit

Get-Process -Name CcmExec -ea SilentlyContinue | Stop-Process -force

Get-service -Name CcmExec | stop-service -Force

Get-service -Name CcmExec | Set-Service -StartupType Disabled