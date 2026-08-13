param($Timer)

Import-Module "$PSScriptRoot/../modules/DnsPostureHelpers.psm1" -Force

try {
    Write-Information 'DomainPostureCollector started.'

    $domains = Get-DomainPostureScope
    if (-not $domains -or $domains.Count -eq 0) {
        Write-Information 'No in-scope domains for DNS posture collection.'
        return
    }

    $rows = [System.Collections.Generic.List[hashtable]]::new()

    foreach ($domain in $domains) {
        $domain = $domain.Trim().ToLowerInvariant()
        if ([string]::IsNullOrWhiteSpace($domain)) { continue }

        $dmarcRow = Get-RecordResult -Domain $domain -RecordName "_dmarc.$domain" -CheckType 'dmarc'
        $rows.Add($dmarcRow)

        $spfRows = Get-SpfResult -Domain $domain
        foreach ($spfRow in $spfRows) {
            $rows.Add($spfRow)
        }

        foreach ($selector in (Get-DkimSelectorsForDomain -Domain $domain)) {
            $dkimRow = Get-RecordResult -Domain $domain -RecordName "$selector._domainkey.$domain" -CheckType 'dkim_selector'
            $rows.Add($dkimRow)
        }

        if (Test-ReportDmarcRelevant -Domain $domain) {
            $reportDmarcRow = Get-RecordResult -Domain $domain -RecordName "_report._dmarc.$domain" -CheckType 'report_dmarc'
            $rows.Add($reportDmarcRow)
        }
    }

    if ($rows.Count -gt 0) {
        Send-DomainPostureToLogAnalytics -Records $rows.ToArray()
    }

    Write-Information "DomainPostureCollector complete. Rows: $($rows.Count)"
}
catch {
    Write-Error "DomainPostureCollector failed: $_"
    Write-Error $_.ScriptStackTrace
    throw
}

if ($Timer.IsPastDue) {
    Write-Warning 'Timer is past due.'
}
