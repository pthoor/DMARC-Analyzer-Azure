# CatchupProcessor - Timer Trigger
# Runs daily to catch any DMARC reports that were missed by the
# real-time Event Grid pipeline (e.g., during outages, subscription gaps,
# or Event Grid delivery failures).
# Checks for unread messages with attachments older than 1 hour.

param($Timer)

Import-Module "$PSScriptRoot/../modules/DmarcHelpers.psm1" -Force

$userId = $env:MAILBOX_USER_ID

if (-not $userId) {
    Write-Error "MAILBOX_USER_ID is not set."
    return
}

Write-Information "Catchup processor started. Checking for unread DMARC reports..."

try {
    $graphToken = Get-ManagedIdentityToken -Resource 'https://graph.microsoft.com'

    # Find unread messages older than 60 minutes (already should have been processed)
    $unreadMessages = Get-UnreadMessages -UserId $userId -OlderThanMinutes 60 -Token $graphToken

    if (-not $unreadMessages -or $unreadMessages.Count -eq 0) {
        Write-Information "No missed DMARC reports found. All caught up."
        return
    }

    Write-Information "Found $($unreadMessages.Count) unread message(s) to process."

    $successCount = 0
    $failCount = 0

    foreach ($message in $unreadMessages) {
        try {
            Write-Information "Processing missed message: $($message.id) - Subject: $($message.subject)"
            Invoke-DmarcReportProcessing -MessageId $message.id
            $successCount++
        }
        catch {
            Write-Warning "Failed to process message $($message.id): $_"
            $failCount++
        }
    }

    Write-Information "Catchup complete. Processed: $successCount, Failed: $failCount"
}
catch {
    Write-Error "Catchup processor failed: $_"
    Write-Error $_.ScriptStackTrace
    throw
}

if ($Timer.IsPastDue) {
    Write-Warning "Catchup timer is past due."
}
