##############
# Desc: Check format
# Author: Can you really call me the author(KVB)? This is lame
# Note: This was abandoned a while ago when I discovered analyzer but I might lean on it for something else later
#############

#Grab your on script
$PS1 = Get-Content ".\Oracle_ODBCInstantClient_12.2.0.1.0_PKG_All.ps1"
$Pattern = "(if|else|elseif|while)(\(|{|}|\))"
[regex]::Matches($PS1, $Pattern).Value
