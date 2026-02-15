#Requires -Version 7.4

<#
.SYNOPSIS
    Shared helper functions for the DMARC-to-Sentinel pipeline.
.DESCRIPTION
    Provides token acquisition (Managed Identity), Microsoft Graph API calls,
    DMARC XML parsing, and Log Analytics ingestion via the Logs Ingestion API.
    All operations use raw REST API calls — no external PowerShell modules.
#>

# ─────────────────────────────────────────────
# Token Acquisition (Managed Identity)
# ─────────────────────────────────────────────

function Get-ManagedIdentityToken {
    <#
    .SYNOPSIS
        Acquires an access token using the Function App's system-assigned Managed Identity.
    .PARAMETER Resource
        The resource URI to request a token for.
        Use 'https://graph.microsoft.com' for Graph API.
        Use 'https://monitor.azure.com' for Logs Ingestion API.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Resource
    )

    $identityEndpoint = $env:IDENTITY_ENDPOINT
    $identityHeader   = $env:IDENTITY_HEADER

    if ([string]::IsNullOrEmpty($identityEndpoint)) {
        $msg = "Managed Identity is not properly configured: environment variable 'IDENTITY_ENDPOINT' is not set or is empty. This is required to acquire a Managed Identity token (resource '$Resource')."
        Write-Error $msg
        throw [System.InvalidOperationException]::new($msg)
    }

    if ([string]::IsNullOrEmpty($identityHeader)) {
        $msg = "Managed Identity is not properly configured: environment variable 'IDENTITY_HEADER' is not set or is empty. This is required to acquire a Managed Identity token (resource '$Resource')."
        Write-Error $msg
        throw [System.InvalidOperationException]::new($msg)
    }

    $tokenUri = "$identityEndpoint?resource=$Resource&api-version=2019-08-01"
    $headers  = @{ 'X-IDENTITY-HEADER' = $identityHeader }
    try {
        $response = Invoke-RestMethod -Uri $tokenUri -Headers $headers -Method Get
        return $response.access_token
    }
    catch {
        Write-Error "Failed to acquire Managed Identity token for resource '$Resource': $_"
        throw
    }
}

# ─────────────────────────────────────────────
# Microsoft Graph API Helpers
# ─────────────────────────────────────────────

function Invoke-GraphRequest {
    <#
    .SYNOPSIS
        Makes an authenticated request to the Microsoft Graph API.
    .PARAMETER Uri
        The full Graph API URI (e.g., https://graph.microsoft.com/v1.0/users/...)
    .PARAMETER Method
        HTTP method (GET, POST, PATCH, DELETE). Default: GET.
    .PARAMETER Body
        Optional request body (will be serialized to JSON).
    .PARAMETER Token
        Optional pre-acquired token. If not provided, acquires one via MI.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Uri,

        [ValidateSet('GET', 'POST', 'PATCH', 'DELETE')]
        [string]$Method = 'GET',

        [object]$Body,

        [string]$Token
    )

    if (-not $Token) {
        $Token = Get-ManagedIdentityToken -Resource 'https://graph.microsoft.com'
    }

    $headers = @{
        'Authorization' = "Bearer $Token"
        'Content-Type'  = 'application/json'
    }

    $params = @{
        Uri     = $Uri
        Method  = $Method
        Headers = $headers
    }

    if ($Body) {
        $params['Body'] = ($Body | ConvertTo-Json -Depth 10)
    }

    try {
        return Invoke-RestMethod @params
    }
    catch {
        $statusCode = if ($_.Exception.Response) { $_.Exception.Response.StatusCode.value__ } else { 'N/A' }
        Write-Error "Graph API request failed [$Method $Uri] - HTTP $statusCode : $_"
        throw
    }
}

function Get-MailMessage {
    <#
    .SYNOPSIS
        Fetches a mail message and its attachments from a mailbox.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserId,

        [Parameter(Mandatory)]
        [string]$MessageId,

        [string]$Token
    )

    $uri = "https://graph.microsoft.com/v1.0/users/$UserId/messages/$MessageId"
    $message = Invoke-GraphRequest -Uri $uri -Token $Token

    # Fetch attachments
    $attachmentsUri = "$uri/attachments"
    $attachments = Invoke-GraphRequest -Uri $attachmentsUri -Token $Token

    return @{
        Message     = $message
        Attachments = $attachments.value
    }
}

function Set-MessageRead {
    <#
    .SYNOPSIS
        Marks a mail message as read.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserId,

        [Parameter(Mandatory)]
        [string]$MessageId,

        [string]$Token
    )

    $uri = "https://graph.microsoft.com/v1.0/users/$UserId/messages/$MessageId"
    Invoke-GraphRequest -Uri $uri -Method PATCH -Body @{ isRead = $true } -Token $Token
}

function Get-UnreadMessages {
    <#
    .SYNOPSIS
        Gets unread messages from the mailbox, optionally filtered by age.
    .PARAMETER OlderThanMinutes
        Only return messages received more than this many minutes ago.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserId,

        [int]$OlderThanMinutes = 60,

        [string]$Token
    )

    $cutoff = (Get-Date).AddMinutes(-$OlderThanMinutes).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    $filter = "isRead eq false and receivedDateTime lt $cutoff and hasAttachments eq true"
    $uri = "https://graph.microsoft.com/v1.0/users/$UserId/messages?`$filter=$filter&`$orderby=receivedDateTime asc&`$top=50&`$select=id,subject,receivedDateTime"

    $result = Invoke-GraphRequest -Uri $uri -Token $Token
    return $result.value
}

# ─────────────────────────────────────────────
# Attachment Extraction
# ─────────────────────────────────────────────

function Expand-DmarcAttachments {
    <#
    .SYNOPSIS
        Extracts DMARC XML content from mail attachments.
    .DESCRIPTION
        Handles .xml, .xml.gz, .gz, and .zip file attachments.
        Returns an array of XML content strings.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [array]$Attachments
    )

    $xmlContents = [System.Collections.Generic.List[string]]::new()

    foreach ($attachment in $Attachments) {
        if ($attachment.'@odata.type' -ne '#microsoft.graph.fileAttachment') {
            Write-Verbose "Skipping non-file attachment: $($attachment.name)"
            continue
        }

        $name = $attachment.name.ToLower()
        $contentBytes = [System.Convert]::FromBase64String($attachment.contentBytes)

        try {
            if ($name.EndsWith('.zip')) {
                $xmlContents.AddRange((Expand-ZipAttachment -ContentBytes $contentBytes))
            }
            elseif ($name.EndsWith('.gz') -or $name.EndsWith('.xml.gz')) {
                $xmlContents.Add((Expand-GzipAttachment -ContentBytes $contentBytes))
            }
            elseif ($name.EndsWith('.xml')) {
                $xmlContents.Add([System.Text.Encoding]::UTF8.GetString($contentBytes))
            }
            else {
                Write-Verbose "Skipping unrecognized attachment: $($attachment.name)"
            }
        }
        catch {
            Write-Warning "Failed to extract attachment '$($attachment.name)': $_"
        }
    }

    return $xmlContents.ToArray()
}

function Expand-ZipAttachment {
    [CmdletBinding()]
    param([byte[]]$ContentBytes)

    $xmlContents = [System.Collections.Generic.List[string]]::new()
    $memStream = [System.IO.MemoryStream]::new($ContentBytes)

    try {
        $archive = [System.IO.Compression.ZipArchive]::new($memStream, [System.IO.Compression.ZipArchiveMode]::Read)

        foreach ($entry in $archive.Entries) {
            $entryName = $entry.Name.ToLower()

            if ($entryName.EndsWith('.xml')) {
                $reader = [System.IO.StreamReader]::new($entry.Open())
                try {
                    $xmlContents.Add($reader.ReadToEnd())
                }
                finally {
                    $reader.Dispose()
                }
            }
            elseif ($entryName.EndsWith('.gz')) {
                # Handle nested .xml.gz inside .zip
                $entryStream = $entry.Open()
                try {
                    $entryMemStream = [System.IO.MemoryStream]::new()
                    try {
                        $entryStream.CopyTo($entryMemStream)
                        $xmlContents.Add((Expand-GzipAttachment -ContentBytes $entryMemStream.ToArray()))
                    }
                    finally {
                        $entryMemStream.Dispose()
                    }
                }
                finally {
                    $entryStream.Dispose()
                }
            }
        }

        $archive.Dispose()
    }
    finally {
        $memStream.Dispose()
    }

    return $xmlContents.ToArray()
}

function Expand-GzipAttachment {
    [CmdletBinding()]
    param([byte[]]$ContentBytes)

    $inputStream = [System.IO.MemoryStream]::new($ContentBytes)
    $gzipStream = [System.IO.Compression.GZipStream]::new($inputStream, [System.IO.Compression.CompressionMode]::Decompress)
    $outputStream = [System.IO.MemoryStream]::new()

    try {
        $gzipStream.CopyTo($outputStream)
        return [System.Text.Encoding]::UTF8.GetString($outputStream.ToArray())
    }
    finally {
        $gzipStream.Dispose()
        $inputStream.Dispose()
        $outputStream.Dispose()
    }
}

# ─────────────────────────────────────────────
# DMARC XML Parsing
# ─────────────────────────────────────────────

function ConvertFrom-DmarcXml {
    <#
    .SYNOPSIS
        Parses DMARC aggregate report XML into flat record objects.
    .DESCRIPTION
        Each <record> element becomes one output object containing:
        - Report metadata (org_name, report_id, date range)
        - Policy published (domain, p, sp, pct, adkim, aspf)
        - Row data (source_ip, count, disposition, dkim/spf evaluation)
        - Identifiers (header_from, envelope_from, envelope_to)
        - Auth results (primary + full JSON arrays)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$XmlContent
    )

    $records = [System.Collections.Generic.List[hashtable]]::new()

    try {
        # Explicitly reject XML with DOCTYPE to prevent DTD-based attacks
        if ($XmlContent -match '<!DOCTYPE') {
            throw "XML contains a DOCTYPE declaration, which is not allowed."
        }

        $xml = [System.Xml.XmlDocument]::new()
        # Ensure no external XML resources are resolved
        $xml.XmlResolver = $null
        $xml.LoadXml($XmlContent)
    }
    catch {
        Write-Warning "Failed to parse DMARC XML: $_"
        return @()
    }

    $feedback = $xml.feedback
    if (-not $feedback) {
        Write-Warning "XML does not contain a <feedback> root element."
        return @()
    }

    # ── Report metadata ──
    $meta = $feedback.report_metadata
    $orgName   = $meta.org_name
    $email     = $meta.email
    $extraContact = $meta.extra_contact_info
    $reportId  = $meta.report_id

    $dateBegin = $null
    $dateEnd   = $null
    if ($meta.date_range) {
        $dateBegin = Convert-EpochToIso -Epoch $meta.date_range.begin
        $dateEnd   = Convert-EpochToIso -Epoch $meta.date_range.end
    }

    # ── Policy published ──
    $policy = $feedback.policy_published
    $domain = $policy.domain
    $policyP   = $policy.p
    $policySp  = $policy.sp
    $policyPct = if ($policy.pct) { [int]$policy.pct } else { 100 }
    $adkim     = $policy.adkim
    $aspf      = $policy.aspf
    $fo        = $policy.fo

    # ── Records ──
    $recordElements = $feedback.record
    if (-not $recordElements) {
        Write-Warning "No <record> elements found in report $reportId"
        return @()
    }

    # Handle single record (PowerShell XML treats single child differently)
    if ($recordElements -isnot [System.Array]) {
        $recordElements = @($recordElements)
    }

    foreach ($rec in $recordElements) {
        $row = $rec.row
        $identifiers = $rec.identifiers
        $authResults = $rec.auth_results

        # ── Primary DKIM result ──
        $dkimResults = @($authResults.dkim)
        $primaryDkim = $null
        $primaryDkimDomain = $null
        $primaryDkimSelector = $null
        $dkimJson = '[]'

        if ($dkimResults.Count -gt 0 -and $dkimResults[0]) {
            $primaryDkim = $dkimResults[0].result
            $primaryDkimDomain = $dkimResults[0].domain
            $primaryDkimSelector = $dkimResults[0].selector

            $dkimArray = foreach ($d in $dkimResults) {
                @{
                    domain       = $d.domain
                    result       = $d.result
                    selector     = $d.selector
                    human_result = $d.human_result
                }
            }
            $dkimJson = ($dkimArray | ConvertTo-Json -Depth 5 -Compress)
            if ($dkimResults.Count -eq 1) { $dkimJson = "[$dkimJson]" }
        }

        # ── Primary SPF result ──
        $spfResults = @($authResults.spf)
        $primarySpf = $null
        $primarySpfDomain = $null
        $primarySpfScope = $null
        $spfJson = '[]'

        if ($spfResults.Count -gt 0 -and $spfResults[0]) {
            $primarySpf = $spfResults[0].result
            $primarySpfDomain = $spfResults[0].domain
            $primarySpfScope = $spfResults[0].scope

            $spfArray = foreach ($s in $spfResults) {
                @{
                    domain = $s.domain
                    result = $s.result
                    scope  = $s.scope
                }
            }
            $spfJson = ($spfArray | ConvertTo-Json -Depth 5 -Compress)
            if ($spfResults.Count -eq 1) { $spfJson = "[$spfJson]" }
        }

        # ── Build flat record ──
        $record = @{
            TimeGenerated                  = [datetime]::UtcNow.ToString('o')
            ReportOrgName                  = $orgName
            ReportEmail                    = $email
            ReportExtraContactInfo         = $extraContact
            ReportId                       = $reportId
            ReportDateRangeBegin           = $dateBegin
            ReportDateRangeEnd             = $dateEnd
            Domain                         = $domain
            PolicyPublished_p              = $policyP
            PolicyPublished_sp             = $policySp
            PolicyPublished_pct            = $policyPct
            PolicyPublished_adkim          = $adkim
            PolicyPublished_aspf           = $aspf
            PolicyPublished_fo             = $fo
            SourceIP                       = $row.source_ip
            MessageCount                   = [int]$row.count
            PolicyEvaluated_disposition    = $row.policy_evaluated.disposition
            PolicyEvaluated_dkim           = $row.policy_evaluated.dkim
            PolicyEvaluated_spf            = $row.policy_evaluated.spf
            PolicyEvaluated_reason_type    = if ($row.policy_evaluated.reason -is [System.Array]) {
                ($row.policy_evaluated.reason | ForEach-Object { $_.type }) -join '; '
            } else { $row.policy_evaluated.reason.type }
            PolicyEvaluated_reason_comment = if ($row.policy_evaluated.reason -is [System.Array]) {
                ($row.policy_evaluated.reason | ForEach-Object { $_.comment }) -join '; '
            } else { $row.policy_evaluated.reason.comment }
            HeaderFrom                     = $identifiers.header_from
            EnvelopeFrom                   = $identifiers.envelope_from
            EnvelopeTo                     = $identifiers.envelope_to
            DkimResult                     = $primaryDkim
            DkimDomain                     = $primaryDkimDomain
            DkimSelector                   = $primaryDkimSelector
            SpfResult                      = $primarySpf
            SpfDomain                      = $primarySpfDomain
            SpfScope                       = $primarySpfScope
            DkimAuthResults                = $dkimJson
            SpfAuthResults                 = $spfJson
        }

        $records.Add($record)
    }

    return $records.ToArray()
}

function Convert-EpochToIso {
    [CmdletBinding()]
    param([string]$Epoch)

    if ([string]::IsNullOrWhiteSpace($Epoch)) { return $null }

    try {
        $epochInt = [long]$Epoch
        return [DateTimeOffset]::FromUnixTimeSeconds($epochInt).UtcDateTime.ToString('o')
    }
    catch {
        Write-Warning "Failed to convert epoch '$Epoch' to ISO: $_"
        return $null
    }
}

# ─────────────────────────────────────────────
# Logs Ingestion API
# ─────────────────────────────────────────────

function Send-DmarcRecordsToLogAnalytics {
    <#
    .SYNOPSIS
        Sends DMARC records to Log Analytics via the Logs Ingestion API (DCR).
    .DESCRIPTION
        Uses Managed Identity to authenticate against https://monitor.azure.com.
        Posts records to the DCR's built-in logsIngestion endpoint.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable[]]$Records
    )

    $dcrEndpoint   = $env:DCR_ENDPOINT
    $dcrImmutableId = $env:DCR_IMMUTABLE_ID
    $streamName    = $env:DCR_STREAM_NAME

    if (-not $dcrEndpoint -or -not $dcrImmutableId -or -not $streamName) {
        throw "Missing DCR configuration. Ensure DCR_ENDPOINT, DCR_IMMUTABLE_ID, and DCR_STREAM_NAME are set."
    }

    $token = Get-ManagedIdentityToken -Resource 'https://monitor.azure.com'

    $uri = "$dcrEndpoint/dataCollectionRules/$dcrImmutableId/streams/${streamName}?api-version=2023-01-01"

    $headers = @{
        'Authorization' = "Bearer $token"
        'Content-Type'  = 'application/json'
    }

    $body = $Records | ConvertTo-Json -Depth 10 -Compress
    # Ensure it's always a JSON array
    if ($Records.Count -eq 1) {
        $body = "[$body]"
    }

    try {
        Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -Body $body
        Write-Information "Successfully sent $($Records.Count) records to Log Analytics."
    }
    catch {
        $statusCode = if ($_.Exception.Response) { $_.Exception.Response.StatusCode.value__ } else { 'N/A' }
        Write-Error "Logs Ingestion API request failed - HTTP $statusCode : $_"
        throw
    }
}

# ─────────────────────────────────────────────
# Orchestration
# ─────────────────────────────────────────────

function Invoke-DmarcReportProcessing {
    <#
    .SYNOPSIS
        End-to-end processing of a single mail message containing DMARC reports.
    .DESCRIPTION
        Fetches the message, extracts attachments, parses DMARC XML,
        sends records to Log Analytics, and marks the message as read.
    .PARAMETER MessageId
        The Graph API message ID.
    .PARAMETER UserId
        The mailbox user ID (object ID of the shared mailbox).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$MessageId,

        [string]$UserId = $env:MAILBOX_USER_ID
    )

    Write-Information "Processing message: $MessageId"

    # Get Graph token once for all operations
    $graphToken = Get-ManagedIdentityToken -Resource 'https://graph.microsoft.com'

    # Fetch message and attachments
    $mail = Get-MailMessage -UserId $UserId -MessageId $MessageId -Token $graphToken
    $subject = $mail.Message.subject
    Write-Information "Message subject: $subject"

    if (-not $mail.Attachments -or $mail.Attachments.Count -eq 0) {
        Write-Warning "Message $MessageId has no attachments. Marking as read."
        Set-MessageRead -UserId $UserId -MessageId $MessageId -Token $graphToken
        return
    }

    # Extract XML from attachments
    $xmlContents = Expand-DmarcAttachments -Attachments $mail.Attachments
    Write-Information "Extracted $($xmlContents.Count) XML content(s) from attachments."

    if ($xmlContents.Count -eq 0) {
        Write-Warning "No DMARC XML found in attachments for message $MessageId"
        Set-MessageRead -UserId $UserId -MessageId $MessageId -Token $graphToken
        return
    }

    # Parse all XML contents
    $allRecords = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($xmlContent in $xmlContents) {
        $parsed = ConvertFrom-DmarcXml -XmlContent $xmlContent
        if ($parsed.Count -gt 0) {
            $allRecords.AddRange($parsed)
        }
    }

    Write-Information "Parsed $($allRecords.Count) total DMARC records."

    if ($allRecords.Count -gt 0) {
        # Send in batches of 500 (API limit guidance)
        $batchSize = 500
        for ($i = 0; $i -lt $allRecords.Count; $i += $batchSize) {
            $batch = $allRecords.GetRange($i, [Math]::Min($batchSize, $allRecords.Count - $i)).ToArray()
            Send-DmarcRecordsToLogAnalytics -Records $batch
            Write-Information "Sent batch: $($batch.Count) records (offset $i)."
        }
    }

    # Mark message as read
    Set-MessageRead -UserId $UserId -MessageId $MessageId -Token $graphToken
    Write-Information "Message $MessageId processed and marked as read."
}

# Export all public functions
Export-ModuleMember -Function @(
    'Get-ManagedIdentityToken'
    'Invoke-GraphRequest'
    'Get-MailMessage'
    'Set-MessageRead'
    'Get-UnreadMessages'
    'Expand-DmarcAttachments'
    'ConvertFrom-DmarcXml'
    'Send-DmarcRecordsToLogAnalytics'
    'Invoke-DmarcReportProcessing'
)
