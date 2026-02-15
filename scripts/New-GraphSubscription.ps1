<#
.SYNOPSIS
    Creates a Microsoft Graph change notification subscription for the DMARC mailbox,
    delivered via Azure Event Grid.

.DESCRIPTION
    This script:
    1. Creates a Graph subscription that watches the shared mailbox for new messages.
    2. Delivers notifications via Event Grid (partner topic), not webhooks.
    3. Saves the subscription ID to the Function App's app settings.

    The subscription must be renewed before it expires (max 4230 minutes / ~2.9 days).
    The RenewGraphSubscription timer function handles this automatically.

.PARAMETER FunctionAppName
    Name of the Azure Function App.

.PARAMETER ResourceGroupName
    Resource group containing the Function App.

.PARAMETER MailboxUserId
    Object ID of the shared mailbox user in Entra ID.

.PARAMETER SubscriptionId
    Azure subscription ID (for Event Grid partner topic).

.PARAMETER GraphClientState
    The client state secret used to validate notifications.
    Must match the GRAPH_CLIENT_STATE app setting.

.EXAMPLE
    .\New-GraphSubscription.ps1 `
        -FunctionAppName 'dmarc-func-abc123' `
        -ResourceGroupName 'rg-dmarc' `
        -MailboxUserId '00000000-0000-0000-0000-000000000000' `
        -SubscriptionId '11111111-1111-1111-1111-111111111111' `
        -GraphClientState 'my-secret-state'

.NOTES
    Prerequisites:
    - The Managed Identity must already have Graph Mail.Read permission
      (run Grant-MIExchangeRBAC.ps1 first).
    - Microsoft Graph must be authorized as an Event Grid partner in the Azure subscription.
      This is a one-time manual step in the Azure Portal:
      Event Grid > Partner Registrations > Authorize "Microsoft Graph API"
    - After this script creates the subscription, a Partner Topic appears in the
      resource group. You must activate it and create an Event Subscription
      pointing to the Function App's DmarcReportProcessor endpoint.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$FunctionAppName,

    [Parameter(Mandatory)]
    [string]$ResourceGroupName,

    [Parameter(Mandatory)]
    [string]$MailboxUserId,

    [Parameter(Mandatory)]
    [string]$SubscriptionId,

    [Parameter(Mandatory)]
    [string]$GraphClientState
)

$ErrorActionPreference = 'Stop'

Write-Host "`n═══════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  DMARC Pipeline — Graph Subscription Setup" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════════`n" -ForegroundColor Cyan

# ─────────────────────────────────────────────
# Step 1: Prerequisite check
# ─────────────────────────────────────────────

Write-Host "[1/4] Checking prerequisites..." -ForegroundColor Yellow

Write-Host @"

  Before continuing, ensure you have completed these manual steps:

  1. Authorize Microsoft Graph as an Event Grid partner:
     Azure Portal > Event Grid > Partner Registrations
     > Authorize "Microsoft Graph API"

  2. Run Grant-MIExchangeRBAC.ps1 to set up MI permissions.

"@ -ForegroundColor Gray

$proceed = Read-Host "Have you completed these steps? (y/n)"
if ($proceed -ne 'y') {
    Write-Host "Please complete the prerequisites first." -ForegroundColor Red
    return
}

# ─────────────────────────────────────────────
# Step 2: Get MI token and create subscription
# ─────────────────────────────────────────────

Write-Host "`n[2/4] Creating Graph change notification subscription..." -ForegroundColor Yellow

# We use the user's Graph context (not MI) since this runs from a workstation
$graphContext = Get-MgContext
if (-not $graphContext) {
    Write-Host "  Connecting to Microsoft Graph..." -ForegroundColor Gray
    Connect-MgGraph -Scopes 'Mail.Read'
}

# Build the Event Grid notification URL
$partnerTopicName = "DmarcPipeline-$FunctionAppName"
$notificationUrl = "EventGrid:?azuresubscriptionid=$SubscriptionId&resourcegroup=$ResourceGroupName&partnertopic=$partnerTopicName&location=$(
    (Get-AzResourceGroup -Name $ResourceGroupName).Location
)"

Write-Host "  Notification URL: $notificationUrl" -ForegroundColor Gray

# Create the subscription
$expirationDateTime = (Get-Date).AddMinutes(4200).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.fffffffZ')

$subscriptionBody = @{
    changeType         = 'created'
    notificationUrl    = $notificationUrl
    resource           = "users/$MailboxUserId/mailFolders('Inbox')/messages"
    expirationDateTime = $expirationDateTime
    clientState        = $GraphClientState
} | ConvertTo-Json

Write-Host "  Creating subscription (expires: $expirationDateTime)..." -ForegroundColor Gray

$response = Invoke-MgGraphRequest -Method POST `
    -Uri 'https://graph.microsoft.com/v1.0/subscriptions' `
    -Body $subscriptionBody `
    -ContentType 'application/json'

$graphSubscriptionId = $response.id

Write-Host "  Subscription created: $graphSubscriptionId" -ForegroundColor Green

# ─────────────────────────────────────────────
# Step 3: Save subscription ID to Function App
# ─────────────────────────────────────────────

Write-Host "`n[3/4] Saving subscription ID to Function App settings..." -ForegroundColor Yellow

$app = Get-AzWebApp -Name $FunctionAppName -ResourceGroupName $ResourceGroupName
$appSettings = @{}
foreach ($setting in $app.SiteConfig.AppSettings) {
    $appSettings[$setting.Name] = $setting.Value
}
$appSettings['GRAPH_SUBSCRIPTION_ID'] = $graphSubscriptionId

Set-AzWebApp -Name $FunctionAppName -ResourceGroupName $ResourceGroupName `
    -AppSettings $appSettings | Out-Null

# Verify that the setting was actually saved
$updatedApp = Get-AzWebApp -Name $FunctionAppName -ResourceGroupName $ResourceGroupName
$updatedAppSettings = @{}
foreach ($setting in $updatedApp.SiteConfig.AppSettings) {
    $updatedAppSettings[$setting.Name] = $setting.Value
}

if ($updatedAppSettings.ContainsKey('GRAPH_SUBSCRIPTION_ID') -and `
    $updatedAppSettings['GRAPH_SUBSCRIPTION_ID'] -eq $graphSubscriptionId) {
    Write-Host "  GRAPH_SUBSCRIPTION_ID saved." -ForegroundColor Green
} else {
    Write-Error "Failed to verify that GRAPH_SUBSCRIPTION_ID was saved to app settings."
    throw "GRAPH_SUBSCRIPTION_ID not found or does not match the created subscription ID."
}
# ─────────────────────────────────────────────
# Step 4: Next steps
# ─────────────────────────────────────────────

Write-Host "`n[4/4] Remaining manual steps:" -ForegroundColor Yellow

Write-Host @"

  A Partner Topic named '$partnerTopicName' should now appear
  in resource group '$ResourceGroupName'. Complete these steps:

  1. Go to Azure Portal > Resource Group > '$ResourceGroupName'
  2. Find the Partner Topic '$partnerTopicName'
  3. Click 'Activate' to activate the partner topic
  4. Create an Event Subscription:
     - Name:     dmarc-report-processor
     - Schema:   Cloud Events Schema v1.0
     - Endpoint: Azure Function > $FunctionAppName > DmarcReportProcessor

"@ -ForegroundColor Gray

Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  Graph subscription ID: $graphSubscriptionId" -ForegroundColor Green
Write-Host "  The RenewGraphSubscription timer will keep it alive." -ForegroundColor Green
Write-Host "═══════════════════════════════════════════════════════`n" -ForegroundColor Cyan
