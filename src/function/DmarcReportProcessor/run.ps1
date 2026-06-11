# DmarcReportProcessor - Event Grid Trigger
# Triggered by Microsoft Graph change notifications delivered via Event Grid.
# Extracts the message ID from the notification, processes the DMARC report,
# and ingests records into Log Analytics.

param($eventGridEvent, $TriggerMetadata)

Import-Module "$PSScriptRoot/../modules/DmarcHelpers.psm1" -Force

Write-Information "Event Grid trigger fired. Event type: $($eventGridEvent.eventType)"

function Get-EventGridMessageId {
    param(
        [Parameter(Mandatory = $true)]
        $EventGridEvent,

        $ResourceData
    )

    $candidates = @()

    if ($ResourceData -and $ResourceData.id) {
        $candidates += [string]$ResourceData.id
    }

    if ($EventGridEvent.data -and $EventGridEvent.data.id) {
        $candidates += [string]$EventGridEvent.data.id
    }

    if ($EventGridEvent.subject -and $EventGridEvent.subject -match '/Messages/([^/?]+)$') {
        $candidates += [string]$Matches[1]
    }

    if ($EventGridEvent.data -and $EventGridEvent.data.resource -and $EventGridEvent.data.resource -match '/messages/([^/?]+)$') {
        $candidates += [string]$Matches[1]
    }

    foreach ($candidate in $candidates) {
        if (-not [string]::IsNullOrWhiteSpace($candidate)) {
            return [System.Uri]::UnescapeDataString($candidate)
        }
    }

    return $null
}

try {
    # Extract resource data from the change notification
    # Event Grid payload from Graph has the resource data in the event body
    $resourceData = $eventGridEvent.data.resourceData
    if (-not $resourceData) {
        # Alternative path: the data might be structured differently
        $resourceData = $eventGridEvent.data
    }

    # Validate client state — defence-in-depth even though Event Grid delivery
    # is internal to Azure. Fail closed: without a configured GRAPH_CLIENT_STATE
    # there is no way to authenticate the notification, so refuse to process it.
    # Throwing lets Event Grid retry the event once the operator fixes the setting.
    $expectedClientState = $env:GRAPH_CLIENT_STATE
    if (-not $expectedClientState -or $expectedClientState -like '@Microsoft.KeyVault*') {
        Write-Error "GRAPH_CLIENT_STATE is not configured (or its Key Vault reference is unresolved). Refusing to process the notification — client state validation is mandatory."
        throw "GRAPH_CLIENT_STATE is not configured. Notification rejected."
    }

    # Try to read clientState from the same flexible structure as resourceData
    $receivedClientState = $eventGridEvent.data.clientState
    if (-not $receivedClientState -and $resourceData) {
        $receivedClientState = $resourceData.clientState
    }

    # Constant-time comparison (case-sensitive) so the secret cannot be probed
    # byte-by-byte via response timing. FixedTimeEquals returns false on length mismatch.
    $expectedBytes = [System.Text.Encoding]::UTF8.GetBytes([string]$expectedClientState)
    $receivedBytes = [System.Text.Encoding]::UTF8.GetBytes([string]$receivedClientState)
    if (-not [System.Security.Cryptography.CryptographicOperations]::FixedTimeEquals($expectedBytes, $receivedBytes)) {
        # Do not throw — a spoofed or stale notification must not be retried.
        Write-Error "Client state mismatch. The received client state does not match the expected value."
        return
    }

    # Extract message ID from the validated notification using fallback paths.
    $messageId = Get-EventGridMessageId -EventGridEvent $eventGridEvent -ResourceData $resourceData
    if (-not $messageId) {
        Write-Error "Could not extract message ID from Event Grid event."
        # Log only non-sensitive identifiers from the Event Grid event.
        # Do not log clientState or other potentially sensitive fields.
        $eventSummary = @{
            Id        = $eventGridEvent.id
            EventType = $eventGridEvent.eventType
            Subject   = $eventGridEvent.subject
        }
        Write-Error ("Event payload summary (sensitive fields redacted): " + ($eventSummary | ConvertTo-Json -Depth 5))
        return
    }

    Write-Information "Processing message ID: $messageId"

    # Process the DMARC report
    Invoke-DmarcReportProcessing -MessageId $messageId

    Write-Information "Successfully processed DMARC report from message: $messageId"
}
catch {
    Write-Error "Failed to process DMARC report: $_"
    Write-Error $_.ScriptStackTrace
    throw  # Re-throw so Event Grid knows to retry
}
