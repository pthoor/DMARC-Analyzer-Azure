# BackfillProcessor - HTTP Trigger (admin auth)
# On-demand import of existing DMARC reports from the mailbox.
# Use after initial deployment or after an outage to process historical emails.
#
# Query parameters:
#   days        - How many days back to look (1-365, default 7)
#   includeRead - Process already-read messages too (true/false, default false)
#
# Example:
#   POST /api/BackfillProcessor?days=7&includeRead=false
#   POST /api/BackfillProcessor?days=14&includeRead=true

param($Request)

Import-Module "$PSScriptRoot/../modules/DmarcHelpers.psm1" -Force

$userId = $env:MAILBOX_USER_ID

if (-not $userId) {
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode  = 500
        Body        = (@{ error = 'MAILBOX_USER_ID is not configured in app settings.' } | ConvertTo-Json)
        ContentType = 'application/json'
    })
    return
}

# Parse query parameters with safe defaults
$days = 7
if ($Request.Query.days) {
    $parsedDays = 0
    if ([int]::TryParse($Request.Query.days, [ref]$parsedDays) -and $parsedDays -ge 1 -and $parsedDays -le 365) {
        $days = $parsedDays
    }
    else {
        Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode  = 400
            Body        = (@{ error = 'Invalid days parameter. Must be an integer between 1 and 365.' } | ConvertTo-Json)
            ContentType = 'application/json'
        })
        return
    }
}

$includeRead = $false
if ($Request.Query.includeRead -eq 'true') {
    $includeRead = $true
}

Write-Information "Backfill started. Looking back $days day(s), includeRead=$includeRead"

try {
    $graphToken = Get-ManagedIdentityToken -Resource 'https://graph.microsoft.com'

    # Query mailbox with paging support
    $messages = Get-MailboxMessages -UserId $userId -Days $days -IncludeRead $includeRead -Token $graphToken

    if (-not $messages -or $messages.Count -eq 0) {
        $msg = "No messages found in the last $days day(s) matching the criteria."
        Write-Information $msg
        Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode  = 200
            Body        = (@{ message = $msg; processed = 0; failed = 0; skipped = 0 } | ConvertTo-Json)
            ContentType = 'application/json'
        })
        return
    }

    Write-Information "Found $($messages.Count) message(s) to process."

    $successCount = 0
    $failCount = 0
    $skippedCount = 0

    foreach ($message in $messages) {
        try {
            Write-Information "Processing message: $($message.id) - Subject: $($message.subject) - Received: $($message.receivedDateTime)"
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

    $summary = "Backfill complete. Days=$days, IncludeRead=$includeRead, Processed=$successCount, Failed=$failCount, Skipped=$skippedCount"
    Write-Information $summary

    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode  = 200
        Body        = (@{
            message   = $summary
            days      = $days
            includeRead = $includeRead
            found     = $messages.Count
            processed = $successCount
            failed    = $failCount
            skipped   = $skippedCount
        } | ConvertTo-Json)
        ContentType = 'application/json'
    })
}
catch {
    Write-Error "Backfill processor failed: $_"
    Write-Error $_.ScriptStackTrace
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode  = 500
        Body        = (@{ error = "Backfill failed: $($_.Exception.Message)" } | ConvertTo-Json)
        ContentType = 'application/json'
    })
}
