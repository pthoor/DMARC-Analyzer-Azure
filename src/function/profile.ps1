# Azure Functions profile.ps1
#
# This profile runs on every cold start of the Function App.
# Use it for one-time initialization that applies to all functions.

# Ensure TLS 1.2+ for all outbound HTTPS requests
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

# Set Information preference so Write-Information messages appear in logs
$InformationPreference = 'Continue'

Write-Information "DMARC-to-Sentinel Function App initialized. PowerShell $($PSVersionTable.PSVersion)"
