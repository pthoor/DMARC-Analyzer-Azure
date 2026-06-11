# SetupHelper - HTTP Trigger (admin auth)
# Called by New-GraphSubscription.ps1 to create a Graph change notification
# subscription using the Function App's Managed Identity.
#
# This function exists because creating a subscription on another user's
# mailbox requires application-level Mail.Read — which only the MI has.
# The setup script (running as the operator) cannot get that permission,
# so it invokes this function instead.
#
# Reads MAILBOX_USER_ID and GRAPH_CLIENT_STATE from app settings (env vars).
# Accepts notificationUrl and expirationDateTime in the request body.
# notificationUrl must use the "EventGrid:" scheme; direct "https" webhook URLs
# are rejected unless the ALLOW_WEBHOOK_NOTIFICATION_URL=true app setting is present
# (they would deliver notifications, including the clientState secret, to an
# arbitrary external endpoint).

param($Request)

Import-Module "$PSScriptRoot/../modules/DmarcHelpers.psm1" -Force

try {
    $body = $Request.Body

    if (-not $body.notificationUrl -or -not $body.expirationDateTime) {
        Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode  = 400
            Body        = (@{ error = 'Missing required fields: notificationUrl, expirationDateTime' } | ConvertTo-Json)
            ContentType = 'application/json'
        })
        return
    }

    # Validate that notificationUrl is a well-formed URL with an accepted scheme.
    # This pipeline delivers Graph notifications via Event Grid, so only the
    # "EventGrid:" scheme is accepted by default. A direct "https" webhook URL would
    # send notifications — including the GRAPH_CLIENT_STATE secret — to an arbitrary
    # external endpoint, so it requires explicit opt-in via the
    # ALLOW_WEBHOOK_NOTIFICATION_URL=true app setting.
    $allowWebhookUrl = $env:ALLOW_WEBHOOK_NOTIFICATION_URL -eq 'true'
    try {
        $parsedUri = [System.Uri]::new([string]$body.notificationUrl)
        $schemeAllowed = ($parsedUri.Scheme -eq 'eventgrid') -or
                         ($allowWebhookUrl -and $parsedUri.Scheme -eq 'https')
        if (-not $schemeAllowed) {
            Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
                StatusCode  = 400
                Body        = (@{ error = 'notificationUrl must use the EventGrid scheme. Direct HTTPS webhook URLs require the ALLOW_WEBHOOK_NOTIFICATION_URL=true app setting.' } | ConvertTo-Json)
                ContentType = 'application/json'
            })
            return
        }
    }
    catch {
        Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode  = 400
            Body        = (@{ error = 'notificationUrl is not a valid URL.' } | ConvertTo-Json)
            ContentType = 'application/json'
        })
        return
    }

    $mailboxUserId = $env:MAILBOX_USER_ID
    $clientState   = $env:GRAPH_CLIENT_STATE

    if (-not $mailboxUserId) {
        Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode  = 500
            Body        = (@{ error = 'MAILBOX_USER_ID is not configured in app settings.' } | ConvertTo-Json)
            ContentType = 'application/json'
        })
        return
    }

    if (-not $clientState) {
        Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode  = 500
            Body        = (@{ error = "GRAPH_CLIENT_STATE is not configured. Every Graph subscription must have a clientState for authenticated notifications." } | ConvertTo-Json)
            ContentType = 'application/json'
        })
        return
    }

    if ($clientState -like '@Microsoft.KeyVault*') {
        Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode  = 500
            Body        = (@{ error = "GRAPH_CLIENT_STATE contains an unresolved Key Vault reference. Ensure the Function App's Managed Identity has been granted 'Key Vault Secrets User' on the vault and that role propagation is complete before calling SetupHelper." } | ConvertTo-Json)
            ContentType = 'application/json'
        })
        return
    }

    Write-Information "Creating Graph subscription for user $mailboxUserId"

    $subscriptionBody = @{
        changeType         = 'created'
        notificationUrl    = $body.notificationUrl
        resource           = "users/$mailboxUserId/mailFolders('Inbox')/messages"
        expirationDateTime = $body.expirationDateTime
    }

    if ($clientState) {
        $subscriptionBody['clientState'] = $clientState
    }

    $graphToken = Get-ManagedIdentityToken -Resource 'https://graph.microsoft.com'
    $result = Invoke-GraphRequest -Uri 'https://graph.microsoft.com/v1.0/subscriptions' `
        -Method POST -Body $subscriptionBody -Token $graphToken

    Write-Information "Subscription created: $($result.id)"

    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode  = 200
        Body        = (@{ subscriptionId = $result.id } | ConvertTo-Json)
        ContentType = 'application/json'
    })
}
catch {
    Write-Error "SetupHelper failed: $_"
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode  = 500
        Body        = (@{ error = 'An internal error occurred. Check the function logs for details.' } | ConvertTo-Json)
        ContentType = 'application/json'
    })
}
