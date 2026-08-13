#Requires -Version 7.4

Import-Module "$PSScriptRoot/DmarcHelpers.psm1" -Force

function Get-DomainPostureScope {
    [CmdletBinding()]
    param()

    $rawScopes = @(
        $env:DOMAIN_POSTURE_SCOPE,
        $env:DMARC_DOMAIN_SCOPE,
        $env:DOMAIN_SCOPE
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    $domains = [System.Collections.Generic.List[string]]::new()
    $seen = [System.Collections.Generic.HashSet[string]]::new()

    foreach ($scopeText in $rawScopes) {
        foreach ($candidate in ($scopeText -split '[;,\r\n]+')) {
            $domain = ($candidate.Trim() | ForEach-Object { $_.TrimEnd('.') }).Trim()
            if ([string]::IsNullOrWhiteSpace($domain)) { continue }
            if ($domain.StartsWith('.')) { $domain = $domain.Substring(1) }
            if ($domain.StartsWith('*') -or $domain -match '^[^a-zA-Z0-9.-]+$') { continue }
            if ($domain -notmatch '^(?=.{1,253}$)(?:[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)(?:\.(?:[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?))*$') { continue }

            $normalized = $domain.ToLowerInvariant()
            if ($seen.Add($normalized)) {
                $domains.Add($normalized)
            }
        }
    }

    return $domains.ToArray()
}

function Get-NslookupTxtRecords {
    <#
    .SYNOPSIS
        Parses `nslookup -type=txt` text output into TXT record objects shaped like
        Resolve-DnsName's output (an array of objects exposing a Strings array).
    .DESCRIPTION
        A single TXT record can be split across multiple quoted character-strings on one
        "text = ..." line (e.g. long DKIM keys/SPF records exceeding 255 bytes). Every quoted
        segment on a line is grouped into that line's own Strings array so the chunks of ONE
        record stay together (and get concatenated downstream, not "; "-joined as if they were
        separate records). A separate line is treated as a separate record.
        Kept as its own function so it can be unit tested without invoking a real nslookup binary.
    #>
    [CmdletBinding()]
    param(
        [AllowNull()]
        [string[]]$RawOutput
    )

    $entries = @()
    foreach ($line in @($RawOutput)) {
        if ($null -eq $line -or $line -notmatch 'text\s*=') { continue }

        $quotedSegments = [regex]::Matches($line, '"([^"]*)"')
        if ($quotedSegments.Count -eq 0) { continue }

        $entries += [pscustomobject]@{
            Strings = @($quotedSegments | ForEach-Object { $_.Groups[1].Value })
        }
    }

    return $entries
}

function Invoke-DnsTxtLookup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Name
    )

    if (Get-Command Resolve-DnsName -ErrorAction SilentlyContinue) {
        return Resolve-DnsName -Name $Name -Type TXT -ErrorAction SilentlyContinue
    }

    $nslookup = Get-Command nslookup -ErrorAction SilentlyContinue
    if ($null -eq $nslookup) {
        # Neither resolver is available on this host. Returning @() here would be
        # indistinguishable downstream from a genuine "no DNS record" result, so
        # every check would silently report false negatives (e.g. "domain has no
        # DMARC record") instead of surfacing the real problem: this host can't
        # resolve DNS at all. Fail loudly instead.
        throw "No DNS TXT resolver is available on this host: neither 'Resolve-DnsName' (Windows) nor 'nslookup' was found. Install a DNS lookup tool (e.g. the 'dnsutils'/'bind-tools' package) in the Function App's Linux runtime before enabling DomainPostureCollector."
    }

    $rawOutput = & $nslookup.Source -type=txt $Name 2>$null
    if ($null -eq $rawOutput) {
        return @()
    }

    return Get-NslookupTxtRecords -RawOutput @($rawOutput)
}

function Get-RecordResult {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Domain,

        [Parameter(Mandatory)]
        [string]$RecordName,

        [Parameter(Mandatory)]
        [string]$CheckType
    )

    $result = Invoke-DnsTxtLookup -Name $RecordName
    $rawEntries = @()
    foreach ($item in @($result)) {
        if ($item.Strings) {
            $rawEntries += ($item.Strings -join ' ')
        }
        elseif ($item -is [string]) {
            $rawEntries += $item
        }
    }

    $recordRaw = (($rawEntries | Where-Object { $_ -ne $null }) | Select-Object -Unique) -join '; '

    $parsed = [ordered]@{}
    if ($CheckType -eq 'dmarc' -and -not [string]::IsNullOrWhiteSpace($recordRaw)) {
        foreach ($part in ($recordRaw -split ';')) {
            $segment = $part.Trim()
            if ([string]::IsNullOrWhiteSpace($segment)) { continue }
            $kv = $segment -split '=', 2
            if ($kv.Count -eq 2) {
                $parsed[$kv[0].Trim().ToLowerInvariant()] = $kv[1].Trim()
            }
        }
    }
    elseif ($CheckType -eq 'dkim_selector' -and -not [string]::IsNullOrWhiteSpace($recordRaw)) {
        $parsed = @{ selector = ([regex]::Match($RecordName, '^(?<selector>[^.]+)\._domainkey\.').Groups['selector'].Value); txt = $recordRaw }
    }
    elseif ($CheckType -eq 'report_dmarc' -and -not [string]::IsNullOrWhiteSpace($recordRaw)) {
        $parsed = @{ report = $recordRaw }
    }

    $status = 'ok'
    if ($null -eq $result -or $result.Count -eq 0) { $status = 'no_record' }
    if ([string]::IsNullOrWhiteSpace($recordRaw)) { $status = 'no_record' }

    return [ordered]@{
        TimeGenerated = [DateTime]::UtcNow
        Domain        = $Domain
        CheckType     = $CheckType
        RecordName    = $RecordName
        RecordRaw     = $recordRaw
        ParsedFields  = ($parsed | ConvertTo-Json -Compress -Depth 10)
        Status        = $status
        CheckedAt     = [DateTime]::UtcNow
    }
}

function Get-SpfResult {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Domain
    )

    $rows = [System.Collections.Generic.List[hashtable]]::new()
    $seen = [System.Collections.Generic.HashSet[string]]::new()

    function Add-SpfRecord {
        param(
            [string]$LookupName,
            [int]$Depth
        )

        if ($Depth -gt 10) {
            $rows.Add([ordered]@{
                TimeGenerated = [DateTime]::UtcNow
                Domain        = $Domain
                CheckType     = 'spf'
                RecordName    = $LookupName
                RecordRaw     = ''
                ParsedFields  = '{"lookup_limit_exceeded":true}'
                Status        = 'lookup_limit_exceeded'
                CheckedAt     = [DateTime]::UtcNow
            })
            return
        }

        if (-not $seen.Add($LookupName.ToLowerInvariant())) { return }

        $dnsResult = Invoke-DnsTxtLookup -Name $LookupName
        $values = @()
        foreach ($item in @($dnsResult)) {
            if ($item.Strings) { $values += ($item.Strings -join ' ') }
            elseif ($item -is [string]) { $values += $item }
        }

        $rawText = (($values | Where-Object { $_ -ne $null }) | Select-Object -Unique) -join '; '
        $status = if ([string]::IsNullOrWhiteSpace($rawText)) { 'no_record' } else { 'ok' }

        $rows.Add([ordered]@{
            TimeGenerated = [DateTime]::UtcNow
            Domain        = $Domain
            CheckType     = 'spf'
            RecordName    = $LookupName
            RecordRaw     = $rawText
            ParsedFields  = (@{ raw = $rawText; depth = $Depth } | ConvertTo-Json -Compress -Depth 10)
            Status        = $status
            CheckedAt     = [DateTime]::UtcNow
        })

        if (-not [string]::IsNullOrWhiteSpace($rawText)) {
            $matches = [regex]::Matches($rawText, 'include:([A-Za-z0-9._-]+)|redirect=([A-Za-z0-9._-]+)|(?:\b(?:a|mx|ptr|exists)\b)', 'IgnoreCase')
            foreach ($match in $matches) {
                $includeTarget = $match.Groups[1].Value
                if ([string]::IsNullOrWhiteSpace($includeTarget)) { $includeTarget = $match.Groups[2].Value }
                if (-not [string]::IsNullOrWhiteSpace($includeTarget)) {
                    Add-SpfRecord -LookupName $includeTarget -Depth ($Depth + 1)
                }
            }
        }
    }

    Add-SpfRecord -LookupName $Domain -Depth 0
    return $rows.ToArray()
}

function Get-DkimSelectorsForDomain {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Domain
    )

    $candidates = [System.Collections.Generic.List[string]]::new()
    $seen = [System.Collections.Generic.HashSet[string]]::new()

    foreach ($source in @(
        $env:DKIM_SELECTORS,
        $env:DOMAIN_POSTURE_DKIM_SELECTORS,
        $env:DMARC_DKIM_SELECTORS
    )) {
        if ([string]::IsNullOrWhiteSpace($source)) { continue }

        foreach ($entry in ($source -split '[;,\r\n]+')) {
            $selector = $entry.Trim().TrimEnd('.')
            if ([string]::IsNullOrWhiteSpace($selector)) { continue }
            $selector = $selector.TrimStart('_')
            if ($selector -match '^(?:[A-Za-z0-9][A-Za-z0-9_-]{0,63})$') {
                $normalized = $selector.ToLowerInvariant()
                if ($seen.Add($normalized)) { $candidates.Add($normalized) }
            }
        }
    }

    if ($candidates.Count -eq 0) {
        $commonDefaults = @('default', 'selector1', 'selector2', 'mail', 'google', 's1', 's2', 'dkim')
        foreach ($selector in $commonDefaults) {
            $normalized = $selector.ToLowerInvariant()
            if ($seen.Add($normalized)) { $candidates.Add($normalized) }
        }
    }

    foreach ($selector in $candidates.ToArray()) {
        $lookup = "$selector._domainkey.$Domain"
        $records = Invoke-DnsTxtLookup -Name $lookup
        if ($null -ne $records -and @($records).Count -gt 0) {
            continue
        }

        $candidates.Remove($selector)
    }

    return $candidates.ToArray()
}

function Test-ReportDmarcRelevant {
    <#
    .SYNOPSIS
        Determines whether a domain's DMARC record authorizes external reporting (rua/ruf).
    .DESCRIPTION
        Takes the already-resolved `_dmarc.<domain>` TXT text (e.g. from Get-RecordResult's
        RecordRaw) rather than re-resolving DNS itself - the caller has typically already
        fetched this exact record for the 'dmarc' CheckType row, and issuing a second,
        independent DNS query here would double DNS I/O per domain and risk the two lookups
        disagreeing under transient DNS flakiness.
    #>
    [CmdletBinding()]
    param(
        [AllowEmptyString()]
        [AllowNull()]
        [string]$DmarcRecordRaw
    )

    if ([string]::IsNullOrWhiteSpace($DmarcRecordRaw)) {
        return $false
    }

    foreach ($segment in ($DmarcRecordRaw -split ';')) {
        $normalized = $segment.Trim()
        if ([string]::IsNullOrWhiteSpace($normalized)) { continue }

        if ($normalized -match '(?i)^(?:rua|ruf)\s*=') {
            return $true
        }
    }

    return $false
}

function Send-DomainPostureToLogAnalytics {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable[]]$Records
    )

    # Deliberately separate from DCR_ENDPOINT/DCR_IMMUTABLE_ID/DCR_STREAM_NAME, which the
    # DMARC report ingestion functions use for the DMARCReports_CL stream. Reusing those
    # here would silently post posture rows (Domain/CheckType/RecordRaw/...) into the
    # DMARC report stream the moment this function shares a Function App with them -
    # this collector has no dedicated DCR/table yet (see docs/BACKLOG.md C0), so it must
    # fail closed rather than guess at a shared, wrong-shaped destination.
    $dcrEndpoint = $env:POSTURE_DCR_ENDPOINT
    $dcrImmutableId = $env:POSTURE_DCR_IMMUTABLE_ID
    $streamName = $env:POSTURE_DCR_STREAM_NAME

    if (-not $dcrEndpoint -or -not $dcrImmutableId -or -not $streamName) {
        throw 'Missing DCR configuration. Ensure POSTURE_DCR_ENDPOINT, POSTURE_DCR_IMMUTABLE_ID, and POSTURE_DCR_STREAM_NAME are set to a dedicated DomainPosture_CL data collection rule/stream - do not point these at the DMARC report ingestion DCR.'
    }

    $token = Get-ManagedIdentityToken -Resource 'https://monitor.azure.com'
    $uri = "$dcrEndpoint/dataCollectionRules/$dcrImmutableId/streams/$([System.Uri]::EscapeDataString($streamName))?api-version=2023-01-01"

    $headers = @{
        'Authorization' = "Bearer $token"
        'Content-Type' = 'application/json'
    }

    $body = $Records | ConvertTo-Json -Depth 10 -Compress
    if ($Records.Count -eq 1) {
        $body = "[$body]"
    }

    try {
        Invoke-WithRetry -ScriptBlock { Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -Body $body } | Out-Null
        Write-Information "Successfully sent $($Records.Count) domain posture record(s) to Log Analytics."
    }
    catch {
        $statusCode = if ($_.Exception.Response) { $_.Exception.Response.StatusCode.value__ } else { 'N/A' }
        Write-Error "Domain posture Logs Ingestion API request failed - HTTP $statusCode : $_"
        throw
    }
}

Export-ModuleMember -Function @(
    'Get-DomainPostureScope',
    'Get-RecordResult',
    'Get-SpfResult',
    'Get-DkimSelectorsForDomain',
    'Test-ReportDmarcRelevant',
    'Send-DomainPostureToLogAnalytics'
)
