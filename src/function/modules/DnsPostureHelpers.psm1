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

    $result = Resolve-DnsName -Name $RecordName -Type TXT -ErrorAction SilentlyContinue
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

        $dnsResult = Resolve-DnsName -Name $LookupName -Type TXT -ErrorAction SilentlyContinue
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
        $records = Resolve-DnsName -Name $lookup -Type TXT -ErrorAction SilentlyContinue
        if ($null -ne $records -and @($records).Count -gt 0) {
            continue
        }

        $candidates.Remove($selector)
    }

    return $candidates.ToArray()
}

function Test-ReportDmarcRelevant {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Domain
    )

    $recordName = "_dmarc.$Domain"
    $records = Resolve-DnsName -Name $recordName -Type TXT -ErrorAction SilentlyContinue
    foreach ($record in @($records)) {
        $entries = @()
        foreach ($item in @($record.Strings)) {
            if ($item) { $entries += $item }
        }
        if ($entries.Count -eq 0 -and $record -is [string]) {
            $entries = @($record)
        }

        foreach ($entry in $entries) {
            $normalized = $entry.Trim()
            if ([string]::IsNullOrWhiteSpace($normalized)) { continue }

            if ($normalized -match '(?i)(?:^|;\s*)(?:rua|ruf)\s*=') {
                return $true
            }
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

    $dcrEndpoint = $env:DCR_ENDPOINT
    $dcrImmutableId = $env:DCR_IMMUTABLE_ID
    $streamName = $env:DCR_STREAM_NAME

    if (-not $dcrEndpoint -or -not $dcrImmutableId -or -not $streamName) {
        throw 'Missing DCR configuration. Ensure DCR_ENDPOINT, DCR_IMMUTABLE_ID, and DCR_STREAM_NAME are set.'
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

    Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -Body $body | Out-Null
}

Export-ModuleMember -Function @(
    'Get-DomainPostureScope',
    'Get-RecordResult',
    'Get-SpfResult',
    'Get-DkimSelectorsForDomain',
    'Test-ReportDmarcRelevant',
    'Send-DomainPostureToLogAnalytics'
)
