# DmarcReportProcessor - Event Grid Trigger
# Triggered by Microsoft Graph change notifications delivered via Event Grid.
# Extracts the message ID from the notification, processes the DMARC report,
# and ingests records into Log Analytics.

param($eventGridEvent, $TriggerMetadata)

Import-Module "$PSScriptRoot/../modules/DmarcHelpers.psm1" -Force

Write-Information "Event Grid trigger fired. Event type: $($eventGridEvent.eventType)"

try {
    # Extract resource data from the change notification
    # Event Grid payload from Graph has the resource data in the event body
    $resourceData = $eventGridEvent.data.resourceData
    if (-not $resourceData) {
        # Alternative path: the data might be structured differently
        $resourceData = $eventGridEvent.data
    }

    # Validate client state if configured - fail fast before any data extraction
    $expectedClientState = $env:GRAPH_CLIENT_STATE
    if ($expectedClientState) {
        # Try to read clientState from the same flexible structure as resourceData
        $receivedClientState = $eventGridEvent.data.clientState
        if (-not $receivedClientState -and $resourceData) {
            $receivedClientState = $resourceData.clientState
        }
        if ($receivedClientState -ne $expectedClientState) {
            Write-Error "Client state mismatch. The received client state does not match the expected value."
            return
        }
    }

    # Extract message ID from the validated notification
    $messageId = $resourceData.id
    if (-not $messageId) {
        Write-Error "Could not extract message ID from Event Grid event."
        Write-Error "Event payload: $($eventGridEvent | ConvertTo-Json -Depth 10)"
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
