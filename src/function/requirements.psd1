# This file lists PowerShell modules that are automatically installed
# when the Function App runtime starts. We intentionally keep this empty —
# runtime operations use raw REST API calls for maximum reliability, so
# no additional modules are required here. Note: setup/deployment scripts
# (e.g., Grant-MIExchangeRBAC.ps1, New-GraphSubscription.ps1) may require
# external modules such as Microsoft.Graph, Az, or Exchange Online PowerShell.

@{
    # No external module dependencies for the Function App runtime; see setup
    # scripts for any separate module prerequisites used during deployment.
}
