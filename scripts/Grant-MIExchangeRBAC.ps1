<#
.SYNOPSIS
    Grants Exchange Online Application RBAC permissions to the Function App's Managed Identity.

.DESCRIPTION
    This script replaces legacy Application Access Policies with the modern
    Application RBAC approach. It:
    1. Assigns Graph API app roles (Mail.Read, Mail.ReadWrite) to the MI service principal.
    2. Creates an Exchange Online service principal pointer for the MI.
    3. Creates a Management Scope to restrict access to the DMARC shared mailbox only.
    4. Assigns Exchange Application RBAC roles scoped to that mailbox.

    After running this script, the Managed Identity can ONLY access the DMARC
    shared mailbox — not any other mailbox in the tenant.

.PARAMETER FunctionAppName
    Name of the Azure Function App (from Bicep output).

.PARAMETER ResourceGroupName
    Resource group containing the Function App.

.PARAMETER MailboxAddress
    Email address of the DMARC shared mailbox (e.g., dmarc-reports@contoso.com).

.EXAMPLE
    .\Grant-MIExchangeRBAC.ps1 -FunctionAppName 'dmarc-func-abc123' `
        -ResourceGroupName 'rg-dmarc' -MailboxAddress 'dmarc-reports@contoso.com'

.NOTES
    Prerequisites:
    - Az.Accounts PowerShell module (for Invoke-AzRestMethod)
    - Microsoft.Graph PowerShell module (for app role assignment)
    - Exchange Online PowerShell module (for Application RBAC)
    - Global Admin or Exchange Admin + Privileged Role Administrator
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$FunctionAppName,

    [Parameter(Mandatory)]
    [string]$ResourceGroupName,

    [Parameter(Mandatory)]
    [string]$MailboxAddress
)

$ErrorActionPreference = 'Stop'

Write-Host "`n═══════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  DMARC Pipeline — Managed Identity Permission Setup" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════════`n" -ForegroundColor Cyan

# ─────────────────────────────────────────────
# Step 1: Get the Managed Identity details
# ─────────────────────────────────────────────

Write-Host "[1/5] Retrieving Managed Identity details..." -ForegroundColor Yellow

# Get the Function App's MI principal ID via ARM REST (avoids Az.Websites dependency)
$subscriptionId = (Get-AzContext).Subscription.Id
$appPath = "/subscriptions/$subscriptionId/resourceGroups/$ResourceGroupName/providers/Microsoft.Web/sites/${FunctionAppName}?api-version=2024-04-01"
$appResponse = Invoke-AzRestMethod -Path $appPath -Method GET
if ($appResponse.StatusCode -ne 200) {
    throw "Function App '$FunctionAppName' not found in resource group '$ResourceGroupName'."
}
$appJson = $appResponse.Content | ConvertFrom-Json
$miPrincipalId = $appJson.identity.principalId

if (-not $miPrincipalId) {
    throw "Function App '$FunctionAppName' does not have a system-assigned Managed Identity enabled."
}

Write-Host "  MI Principal ID : $miPrincipalId" -ForegroundColor Gray

# Connect to Microsoft Graph (needed for service principal lookup and role assignment)
$graphContext = Get-MgContext
if (-not $graphContext) {
    Write-Host "  Connecting to Microsoft Graph..." -ForegroundColor Gray
    Connect-MgGraph -Scopes 'AppRoleAssignment.ReadWrite.All'
}

# Get the full service principal details via Microsoft Graph SDK (need AppId)
$miSp = Get-MgServicePrincipal -Filter "id eq '$miPrincipalId'"
$miAppId = $miSp.AppId
$miObjectId = $miSp.Id

Write-Host "  MI App ID       : $miAppId" -ForegroundColor Gray
Write-Host "  MI Object ID    : $miObjectId" -ForegroundColor Gray

# ─────────────────────────────────────────────
# Step 2: Assign Graph API app roles to the MI
# ─────────────────────────────────────────────

Write-Host "`n[2/5] Assigning Microsoft Graph app roles to Managed Identity..." -ForegroundColor Yellow

$graphSpId = (Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'").Id
$graphSp = Get-MgServicePrincipal -ServicePrincipalId $graphSpId

$requiredRoles = @('Mail.Read', 'Mail.ReadWrite')

foreach ($roleName in $requiredRoles) {
    $role = $graphSp.AppRoles | Where-Object { $_.Value -eq $roleName }
    if (-not $role) {
        Write-Warning "  Could not find app role '$roleName' on Microsoft Graph."
        continue
    }

    # Check if already assigned
    $existing = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $miObjectId |
        Where-Object { $_.AppRoleId -eq $role.Id }

    if ($existing) {
        Write-Host "  $roleName — already assigned." -ForegroundColor Gray
    }
    else {
        New-MgServicePrincipalAppRoleAssignment `
            -ServicePrincipalId $miObjectId `
            -PrincipalId $miObjectId `
            -ResourceId $graphSpId `
            -AppRoleId $role.Id | Out-Null
        Write-Host "  $roleName — assigned." -ForegroundColor Green
    }
}

# ─────────────────────────────────────────────
# Step 3: Create Exchange Online service principal
# ─────────────────────────────────────────────

Write-Host "`n[3/5] Creating Exchange Online service principal pointer..." -ForegroundColor Yellow

# Connect to Exchange Online (if not already)
try {
    Get-OrganizationConfig | Out-Null
}
catch {
    Write-Host "  Connecting to Exchange Online..." -ForegroundColor Gray
    # Use -Device (device-code flow) to avoid WAM broker, which is unavailable
    # on Linux and can conflict with the MSAL version loaded by Microsoft.Graph.
    Connect-ExchangeOnline -Device -ShowBanner:$false
}

# Check if already exists
$existingExoSp = Get-ServicePrincipal -Identity $miAppId -ErrorAction SilentlyContinue
if ($existingExoSp) {
    Write-Host "  Exchange service principal already exists." -ForegroundColor Gray
}
else {
    New-ServicePrincipal -AppId $miAppId -ObjectId $miObjectId -DisplayName "DMARC Pipeline ($FunctionAppName)"
    Write-Host "  Exchange service principal created." -ForegroundColor Green
}

# ─────────────────────────────────────────────
# Step 4: Create Management Scope for the DMARC mailbox
# ─────────────────────────────────────────────

Write-Host "`n[4/5] Creating Management Scope for mailbox restriction..." -ForegroundColor Yellow

try {
    $parsedMailboxAddress = [System.Net.Mail.MailAddress]::new($MailboxAddress)
} catch {
    throw "MailboxAddress '$MailboxAddress' is not a valid email address."
}

$scopeName = "DMARC-Mailbox-$($parsedMailboxAddress.User)"
$existingScope = Get-ManagementScope -Identity $scopeName -ErrorAction SilentlyContinue
if ($existingScope) {
    Write-Host "  Management Scope '$scopeName' already exists." -ForegroundColor Gray
}
else {
    # Scope the MI to only the DMARC shared mailbox
    New-ManagementScope -Name $scopeName `
        -RecipientRestrictionFilter "PrimarySmtpAddress -eq '$MailboxAddress'"
    Write-Host "  Management Scope '$scopeName' created." -ForegroundColor Green
}

# ─────────────────────────────────────────────
# Step 5: Assign Application RBAC roles
# ─────────────────────────────────────────────

Write-Host "`n[5/5] Assigning Exchange Application RBAC roles..." -ForegroundColor Yellow

$exoRoles = @('Application Mail.Read', 'Application Mail.ReadWrite')

foreach ($roleName in $exoRoles) {
    $assignmentName = "$roleName - DMARC Pipeline"

    $existing = Get-ManagementRoleAssignment -Identity $assignmentName -ErrorAction SilentlyContinue
    if ($existing) {
        if (($existing.App -eq $miAppId) -and ($existing.CustomResourceScope -eq $scopeName)) {
            Write-Host "  $roleName — already assigned." -ForegroundColor Gray
        }
        else {
            Write-Host "  $roleName — existing assignment '$assignmentName' has different properties. Recreating with correct App and scope..." -ForegroundColor Yellow
            Remove-ManagementRoleAssignment -Identity $assignmentName -Confirm:$false
            New-ManagementRoleAssignment `
                -Name $assignmentName `
                -Role $roleName `
                -App $miAppId `
                -CustomResourceScope $scopeName
            Write-Host "  $roleName — reassigned with scope '$scopeName'." -ForegroundColor Green
        }
    }
    else {
        New-ManagementRoleAssignment `
            -Name $assignmentName `
            -Role $roleName `
            -App $miAppId `
            -CustomResourceScope $scopeName
        Write-Host "  $roleName — assigned with scope '$scopeName'." -ForegroundColor Green
    }
}

# ─────────────────────────────────────────────
# Summary
# ─────────────────────────────────────────────

Write-Host "`n═══════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  Setup complete!" -ForegroundColor Green
Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "`n  The Managed Identity for '$FunctionAppName' now has:"
Write-Host "  • Graph API: Mail.Read + Mail.ReadWrite (app roles)"
Write-Host "  • Exchange RBAC: Scoped to '$MailboxAddress' only"
Write-Host "`n  Note: Permission changes may take 30 min to 2 hours to propagate."
Write-Host "  Use Test-ServicePrincipalAuthorization to verify:`n"
Write-Host "  Test-ServicePrincipalAuthorization -Identity '$miAppId' -Resource '$MailboxAddress' | Format-Table`n" -ForegroundColor Gray

# Important: Do NOT remove the Graph app role consents in Entra ID.
# The Exchange Application RBAC scoping works _in addition_ to the Graph roles.
# The Graph roles provide the permission; the RBAC scope restricts where they apply.
# However, you should ensure there are no BROAD unscoped Mail.Read/Mail.ReadWrite
# consents in Entra ID → Enterprise Applications → your MI → Permissions.
# The app role assignments we made above are to the MI specifically, so they're fine.
