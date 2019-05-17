Get-service -Name CcmExec | stop-service -Force

#Need to dig into the registry to make this a delay start...prolly shoul dmake a function for that.
Get-service -Name CcmExec | Set-Service -StartupType Automatic

#when I stop being lazt I cangrab the actual filepath and start the prtocess.
#Get-Process -Name CcmExec -ea SilentlyContinue | Start-Process