# CatchupProcessor - Timer Trigger
# Runs daily to catch any DMARC reports that were missed by the
# real-time Event Grid pipeline (e.g., during outages, subscription gaps,
# or Event Grid delivery failures).
# Checks for unread messages with attachments from the last 2 days.
# Uses paged Graph queries — no longer capped at 50 messages.

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

    # Find unread messages with attachments from the last 2 days (paged, no 50-message cap)
    $unreadMessages = Get-MailboxMessages -UserId $userId -Days 2 -IncludeRead $false -Token $graphToken

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
            $errorRecord = $_
            Write-Warning "Failed to process message $($message.id): $errorRecord"
            if ($errorRecord -and $errorRecord.Exception) {
                Write-Error ("Detailed failure for message {0}: {1}" -f $message.id, $errorRecord.Exception.Message)
            }
            if ($errorRecord -and $errorRecord.ScriptStackTrace) {
                Write-Error $errorRecord.ScriptStackTrace
            }
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
