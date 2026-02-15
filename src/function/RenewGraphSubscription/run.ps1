# RenewGraphSubscription - Timer Trigger
# Renews the Microsoft Graph change notification subscription.
# Outlook message subscriptions expire after max 4230 minutes (~2.9 days).
# This timer runs every 2 hours to keep the subscription alive.

param($Timer)

Import-Module "$PSScriptRoot/../modules/DmarcHelpers.psm1" -Force

$subscriptionId = $env:GRAPH_SUBSCRIPTION_ID

if (-not $subscriptionId) {
    Write-Error "GRAPH_SUBSCRIPTION_ID is not set. Run New-GraphSubscription.ps1 first."
    return
}

Write-Information "Renewing Graph subscription: $subscriptionId"

try {
    $graphToken = Get-ManagedIdentityToken -Resource 'https://graph.microsoft.com'

    # Set new expiration to 4200 minutes from now (just under the 4230 max)
    $newExpiration = (Get-Date).AddMinutes(4200).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.fffffffZ')

    $uri = "https://graph.microsoft.com/v1.0/subscriptions/$subscriptionId"
    $body = @{
        expirationDateTime = $newExpiration
    }

    $result = Invoke-GraphRequest -Uri $uri -Method PATCH -Body $body -Token $graphToken

    Write-Information "Subscription renewed. New expiration: $($result.expirationDateTime)"
}
catch {
    $statusCode = if ($_.Exception.Response) { $_.Exception.Response.StatusCode.value__ } else { $null }

    if ($statusCode -eq 404) {
        Write-Error "Subscription $subscriptionId not found. It may have expired. Re-run New-GraphSubscription.ps1 to create a new one."
    }
    else {
        Write-Error "Failed to renew subscription: $_"
        Write-Error $_.ScriptStackTrace
        throw  # Re-throw for retry
    }
}

if ($Timer.IsPastDue) {
    Write-Warning "Timer is past due — subscription renewal was delayed."
}
