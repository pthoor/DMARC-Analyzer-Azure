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
                Mock Resolve-DnsName {
                    [pscustomobject]@{ Strings = @('v=DKIM1; k=rsa; p=ABC123') }
                } -ParameterFilter { $Name -eq 'default._domainkey.example.com' -or $Name -eq 'selector2._domainkey.example.com' }

                $result = Get-DkimSelectorsForDomain -Domain 'example.com'
                $result | Should -Contain 'default'
                $result | Should -Contain 'selector2'
            }
        }

    }

    Context 'Test-ReportDmarcRelevant' {
        It 'returns true when the DMARC record includes a report URI' {
            InModuleScope DnsPostureHelpers {
                Mock Resolve-DnsName {
                    [pscustomobject]@{ Strings = @('v=DMARC1; p=none; rua=mailto:reports@example.com; ruf=mailto:forensics@example.com') }
                } -ParameterFilter { $Name -eq '_dmarc.example.com' }

                $result = Test-ReportDmarcRelevant -Domain 'example.com'
                $result | Should -BeTrue
            }
        }

        It 'returns false when the DMARC record has no report target' {
            InModuleScope DnsPostureHelpers {
                Mock Resolve-DnsName {
                    [pscustomobject]@{ Strings = @('v=DMARC1; p=reject; adkim=s; aspf=s') }
                } -ParameterFilter { $Name -eq '_dmarc.example.com' }

                $result = Test-ReportDmarcRelevant -Domain 'example.com'
                $result | Should -BeFalse
            }
        }
    }
}
