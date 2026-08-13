#Requires -Version 7.4
#Requires -Modules Pester

BeforeAll {
    Import-Module "$PSScriptRoot/../src/function/modules/DnsPostureHelpers.psm1" -Force
}

Describe 'DnsPostureHelpers Module' {
    BeforeEach {
        $env:DOMAIN_POSTURE_SCOPE = $null
        $env:DMARC_DOMAIN_SCOPE = $null
        $env:DOMAIN_SCOPE = $null
        $env:DKIM_SELECTORS = $null
        $env:DOMAIN_POSTURE_DKIM_SELECTORS = $null
        $env:DMARC_DKIM_SELECTORS = $null
    }

    Context 'Get-DomainPostureScope' {
        It 'normalizes a mixed-scoping list and strips blanks' {
            $env:DOMAIN_POSTURE_SCOPE = 'Example.com; mail.example.com, news.example.com ; ;'

            InModuleScope DnsPostureHelpers {
                $result = Get-DomainPostureScope
                $result.Count | Should -Be 3
                $result[0] | Should -Be 'example.com'
                $result[1] | Should -Be 'mail.example.com'
                $result[2] | Should -Be 'news.example.com'
            }
        }

        It 'ignores malformed values and duplicate domains' {
            $env:DOMAIN_POSTURE_SCOPE = 'Example.com, example.com, .sub.example.com, bad_domain, http://bad.example'

            InModuleScope DnsPostureHelpers {
                $result = Get-DomainPostureScope
                $result | Should -Contain 'example.com'
                $result | Should -Contain 'sub.example.com'
                $result | Should -Not -Contain 'bad_domain'
                $result | Should -Not -Contain 'http://bad.example'
            }
        }
    }

    Context 'Get-DkimSelectorsForDomain' {
        It 'prefers explicitly configured selectors and verifies they resolve' {
            $env:DKIM_SELECTORS = 'Default; selector2, selector2'

            InModuleScope DnsPostureHelpers {
                Mock Invoke-DnsTxtLookup {
                    [pscustomobject]@{ Strings = @('v=DKIM1; k=rsa; p=ABC123') }
                } -ParameterFilter { $Name -eq 'default._domainkey.example.com' -or $Name -eq 'selector2._domainkey.example.com' }

                $result = Get-DkimSelectorsForDomain -Domain 'example.com'
                $result | Should -Contain 'default'
                $result | Should -Contain 'selector2'
            }
        }

    }

    Context 'Test-ReportDmarcRelevant' {
        # Takes the already-fetched DMARC RecordRaw text directly rather than re-resolving
        # DNS itself - the caller (DomainPostureCollector) already fetched this exact record
        # once via Get-RecordResult, so a second independent DNS lookup here would double DNS
        # I/O per domain and risk the two lookups disagreeing under transient DNS flakiness.
        It 'returns true when the DMARC record includes a report URI' {
            InModuleScope DnsPostureHelpers {
                $result = Test-ReportDmarcRelevant -DmarcRecordRaw 'v=DMARC1; p=none; rua=mailto:reports@example.com; ruf=mailto:forensics@example.com'
                $result | Should -BeTrue
            }
        }

        It 'returns false when the DMARC record has no report target' {
            InModuleScope DnsPostureHelpers {
                $result = Test-ReportDmarcRelevant -DmarcRecordRaw 'v=DMARC1; p=reject; adkim=s; aspf=s'
                $result | Should -BeFalse
            }
        }

        It 'returns false when there is no DMARC record at all' {
            InModuleScope DnsPostureHelpers {
                Test-ReportDmarcRelevant -DmarcRecordRaw '' | Should -BeFalse
                Test-ReportDmarcRelevant -DmarcRecordRaw $null | Should -BeFalse
            }
        }
    }

    Context 'Invoke-DnsTxtLookup' {
        It 'throws instead of returning an empty result when neither Resolve-DnsName nor nslookup is available' {
            InModuleScope DnsPostureHelpers {
                Mock Get-Command { $null } -ParameterFilter { $Name -eq 'Resolve-DnsName' -or $Name -eq 'nslookup' }

                { Invoke-DnsTxtLookup -Name '_dmarc.example.com' } | Should -Throw '*No DNS TXT resolver*'
            }
        }

    }

    Context 'Get-NslookupTxtRecords' {
        It 'groups multiple quoted character-strings on one line into a single record (long DKIM key split across chunks)' {
            InModuleScope DnsPostureHelpers {
                $rawOutput = @(
                    'Server:  127.0.0.53',
                    'Address: 127.0.0.53#53',
                    '',
                    'default._domainkey.example.com    text = "v=DKIM1; k=rsa; p=ABC" "123DEF"'
                )

                $result = Get-NslookupTxtRecords -RawOutput $rawOutput

                $result.Count | Should -Be 1
                $result[0].Strings | Should -Be @('v=DKIM1; k=rsa; p=ABC', '123DEF')
            }
        }

        It 'treats separate "text =" lines as separate records' {
            InModuleScope DnsPostureHelpers {
                $rawOutput = @(
                    'example.com    text = "v=spf1 include:one.example.com -all"',
                    'example.com    text = "some-other-unrelated-txt-record"'
                )

                $result = Get-NslookupTxtRecords -RawOutput $rawOutput

                $result.Count | Should -Be 2
                $result[0].Strings | Should -Be @('v=spf1 include:one.example.com -all')
                $result[1].Strings | Should -Be @('some-other-unrelated-txt-record')
            }
        }

        It 'returns an empty array when there is no "text =" line' {
            InModuleScope DnsPostureHelpers {
                $result = Get-NslookupTxtRecords -RawOutput @('Server:  127.0.0.53', '** server can''t find example.com: NXDOMAIN')
                $result.Count | Should -Be 0
            }
        }
    }

    Context 'Send-DomainPostureToLogAnalytics' {
        BeforeEach {
            $env:POSTURE_DCR_ENDPOINT = $null
            $env:POSTURE_DCR_IMMUTABLE_ID = $null
            $env:POSTURE_DCR_STREAM_NAME = $null
            $env:DCR_ENDPOINT = $null
            $env:DCR_IMMUTABLE_ID = $null
            $env:DCR_STREAM_NAME = $null
        }

        It 'throws when POSTURE_DCR_* settings are missing, even if the DMARC ingestion DCR_* settings are set' {
            $env:DCR_ENDPOINT = 'https://dmarc-dce.example.com'
            $env:DCR_IMMUTABLE_ID = 'dcr-dmarc-immutable-id'
            $env:DCR_STREAM_NAME = 'Custom-DMARCReports_CL'

            InModuleScope DnsPostureHelpers {
                { Send-DomainPostureToLogAnalytics -Records @(@{ Domain = 'example.com' }) } | Should -Throw '*POSTURE_DCR_*'
            }
        }

        It 'posts to the dedicated posture stream and never the DMARC ingestion stream' {
            $env:DCR_ENDPOINT = 'https://dmarc-dce.example.com'
            $env:DCR_IMMUTABLE_ID = 'dcr-dmarc-immutable-id'
            $env:DCR_STREAM_NAME = 'Custom-DMARCReports_CL'
            $env:POSTURE_DCR_ENDPOINT = 'https://posture-dce.example.com'
            $env:POSTURE_DCR_IMMUTABLE_ID = 'dcr-posture-immutable-id'
            $env:POSTURE_DCR_STREAM_NAME = 'Custom-DomainPosture_CL'

            InModuleScope DnsPostureHelpers {
                Mock Get-ManagedIdentityToken { 'fake-token' }
                Mock Invoke-RestMethod { }

                Send-DomainPostureToLogAnalytics -Records @(@{ Domain = 'example.com' })

                Should -Invoke Invoke-RestMethod -ParameterFilter {
                    $Uri -like 'https://posture-dce.example.com/dataCollectionRules/dcr-posture-immutable-id/streams/Custom-DomainPosture_CL*'
                }
            }
        }
    }
}
