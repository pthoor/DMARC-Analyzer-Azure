#Requires -Version 7.4
#Requires -Modules Pester

<#
.SYNOPSIS
    Pester tests for DmarcHelpers PowerShell module.
.DESCRIPTION
    Tests all functions in the DmarcHelpers.psm1 module including:
    - Token acquisition
    - Graph API helpers
    - DMARC attachment extraction
    - DMARC XML parsing
    - Log Analytics ingestion
#>

BeforeAll {
    # Import the module
    Import-Module "$PSScriptRoot/../src/function/modules/DmarcHelpers.psm1" -Force
}

Describe 'DmarcHelpers Module' {
    Context 'Module Import' {
        It 'Should import the module successfully' {
            Get-Module DmarcHelpers | Should -Not -BeNullOrEmpty
        }

        It 'Should export expected functions' {
            $exportedFunctions = (Get-Module DmarcHelpers).ExportedFunctions.Keys
            $exportedFunctions | Should -Contain 'Get-ManagedIdentityToken'
            $exportedFunctions | Should -Contain 'Invoke-GraphRequest'
            $exportedFunctions | Should -Contain 'Get-MailMessage'
            $exportedFunctions | Should -Contain 'Set-MessageRead'
            $exportedFunctions | Should -Contain 'Get-MailboxMessages'
            $exportedFunctions | Should -Contain 'Expand-DmarcAttachments'
            $exportedFunctions | Should -Contain 'Get-DomainIdentity'
            $exportedFunctions | Should -Contain 'ConvertFrom-DmarcXml'
            $exportedFunctions | Should -Contain 'ConvertTo-SafeLogText'
            $exportedFunctions | Should -Contain 'Send-DmarcRecordsToLogAnalytics'
            $exportedFunctions | Should -Contain 'Invoke-DmarcReportProcessing'
        }
    }

    Context 'Get-ManagedIdentityToken' {
        It 'Should require IDENTITY_ENDPOINT environment variable' {
            $env:IDENTITY_ENDPOINT = $null
            $env:IDENTITY_HEADER = 'test'
            { Get-ManagedIdentityToken -Resource 'https://graph.microsoft.com' } | Should -Throw '*IDENTITY_ENDPOINT*'
        }

        It 'Should require IDENTITY_HEADER environment variable' {
            $env:IDENTITY_ENDPOINT = 'https://test.endpoint'
            $env:IDENTITY_HEADER = $null
            { Get-ManagedIdentityToken -Resource 'https://graph.microsoft.com' } | Should -Throw '*IDENTITY_HEADER*'
        }
    }

    Context 'Get-DomainIdentity' {
        It 'Should normalize subdomains to registrable domains' {
            $identity = Get-DomainIdentity -DomainName 'mail.example.com'
            $identity.BaseDomain | Should -Be 'example.com'
            $identity.OrgDomain | Should -Be 'example.com'
            $identity.IsSubdomain | Should -Be $true
        }

        It 'Should support multi-part public suffixes like .co.uk' {
            $identity = Get-DomainIdentity -DomainName 'mail.foo.example.co.uk'
            $identity.BaseDomain | Should -Be 'example.co.uk'
            $identity.OrgDomain | Should -Be 'example.co.uk'
            $identity.IsSubdomain | Should -Be $true
        }

        It 'Should canonicalize domain casing and trailing dots before deriving identity' {
            $identity = Get-DomainIdentity -DomainName ' MAIL.FOO.EXAMPLE.CO.UK. '
            $identity.BaseDomain | Should -Be 'example.co.uk'
            $identity.OrgDomain | Should -Be 'example.co.uk'
            $identity.IsSubdomain | Should -Be $true
        }
    }

    Context 'ConvertFrom-DmarcXml' {
        It 'Should parse valid DMARC XML' {
            $validXml = @'
<?xml version="1.0" encoding="UTF-8"?>
<feedback>
  <report_metadata>
    <org_name>google.com</org_name>
    <email>noreply-dmarc-support@google.com</email>
    <report_id>12345678901234567890</report_id>
    <date_range>
      <begin>1704067200</begin>
      <end>1704153599</end>
    </date_range>
  </report_metadata>
  <policy_published>
    <domain>example.com</domain>
    <adkim>r</adkim>
    <aspf>r</aspf>
    <p>none</p>
    <sp>none</sp>
    <pct>100</pct>
  </policy_published>
  <record>
    <row>
      <source_ip>192.0.2.1</source_ip>
      <count>5</count>
      <policy_evaluated>
        <disposition>none</disposition>
        <dkim>pass</dkim>
        <spf>pass</spf>
      </policy_evaluated>
    </row>
    <identifiers>
      <header_from>example.com</header_from>
      <envelope_from>example.com</envelope_from>
    </identifiers>
    <auth_results>
      <dkim>
        <domain>example.com</domain>
        <result>pass</result>
        <selector>default</selector>
      </dkim>
      <spf>
        <domain>example.com</domain>
        <result>pass</result>
        <scope>mfrom</scope>
      </spf>
    </auth_results>
  </record>
</feedback>
'@

            $result = ConvertFrom-DmarcXml -XmlContent $validXml
            $result | Should -Not -BeNullOrEmpty
            # When there's one record, it returns a single hashtable, not an array
            if ($result -is [array]) {
                $result.Length | Should -Be 1
                $record = $result[0]
            } else {
                $record = $result
            }
            $record.ReportOrgName | Should -Be 'google.com'
            $record.Domain | Should -Be 'example.com'
            $record.SourceIP | Should -Be '192.0.2.1'
            $record.MessageCount | Should -Be 5
            $record.PolicyEvaluated_dkim | Should -Be 'pass'
            $record.PolicyEvaluated_spf | Should -Be 'pass'
            $record.DkimResult | Should -Be 'pass'
            $record.SpfResult | Should -Be 'pass'
        }

        It 'Should emit duplicate telemetry fields when processing metadata is provided' {
            $validXml = @'
<?xml version="1.0" encoding="UTF-8"?>
<feedback>
  <report_metadata>
    <org_name>google.com</org_name>
    <email>noreply-dmarc-support@google.com</email>
    <report_id>12345678901234567890</report_id>
    <date_range>
      <begin>1704067200</begin>
      <end>1704153599</end>
    </date_range>
  </report_metadata>
  <policy_published>
    <domain>example.com</domain>
    <adkim>r</adkim>
    <aspf>r</aspf>
    <p>none</p>
    <sp>none</sp>
    <pct>100</pct>
  </policy_published>
  <record>
    <row>
      <source_ip>192.0.2.1</source_ip>
      <count>5</count>
      <policy_evaluated>
        <disposition>none</disposition>
        <dkim>pass</dkim>
        <spf>pass</spf>
      </policy_evaluated>
    </row>
    <identifiers>
      <header_from>example.com</header_from>
      <envelope_from>example.com</envelope_from>
    </identifiers>
    <auth_results>
      <dkim>
        <domain>example.com</domain>
        <result>pass</result>
      </dkim>
      <spf>
        <domain>example.com</domain>
        <result>pass</result>
      </spf>
    </auth_results>
  </record>
</feedback>
'@

            $result = ConvertFrom-DmarcXml -XmlContent $validXml -SourceMessageId 'msg-123' -IngestionRunId 'run-abc'
            $record = if ($result -is [array]) { $result[0] } else { $result }

            $record.SourceMessageId | Should -Be 'msg-123'
            $record.IngestionRunId | Should -Be 'run-abc'
            $record.DuplicateTelemetryKey | Should -BeLike 'google.com|12345678901234567890|example.com|*|*'
        }

        It 'Should parse XML with multiple records' {
            $multiRecordXml = @'
<?xml version="1.0" encoding="UTF-8"?>
<feedback>
  <report_metadata>
    <org_name>test.com</org_name>
    <email>test@test.com</email>
    <report_id>123</report_id>
    <date_range>
      <begin>1704067200</begin>
      <end>1704153599</end>
    </date_range>
  </report_metadata>
  <policy_published>
    <domain>example.com</domain>
    <p>none</p>
    <pct>100</pct>
  </policy_published>
  <record>
    <row>
      <source_ip>192.0.2.1</source_ip>
      <count>10</count>
      <policy_evaluated>
        <disposition>none</disposition>
        <dkim>pass</dkim>
        <spf>pass</spf>
      </policy_evaluated>
    </row>
    <identifiers>
      <header_from>example.com</header_from>
    </identifiers>
    <auth_results>
      <dkim>
        <domain>example.com</domain>
        <result>pass</result>
      </dkim>
      <spf>
        <domain>example.com</domain>
        <result>pass</result>
      </spf>
    </auth_results>
  </record>
  <record>
    <row>
      <source_ip>192.0.2.2</source_ip>
      <count>5</count>
      <policy_evaluated>
        <disposition>none</disposition>
        <dkim>fail</dkim>
        <spf>fail</spf>
      </policy_evaluated>
    </row>
    <identifiers>
      <header_from>example.com</header_from>
    </identifiers>
    <auth_results>
      <dkim>
        <domain>example.com</domain>
        <result>fail</result>
      </dkim>
      <spf>
        <domain>example.com</domain>
        <result>fail</result>
      </spf>
    </auth_results>
  </record>
</feedback>
'@

            $result = ConvertFrom-DmarcXml -XmlContent $multiRecordXml
            $result | Should -Not -BeNullOrEmpty
            # Multiple records should return an array
            $result -is [array] | Should -Be $true
            $result.Length | Should -Be 2
            $result[0].SourceIP | Should -Be '192.0.2.1'
            $result[1].SourceIP | Should -Be '192.0.2.2'
        }

        It 'Should handle invalid XML gracefully' {
            $invalidXml = '<invalid>xml</not-closed>'
            $result = ConvertFrom-DmarcXml -XmlContent $invalidXml
            $result | Should -BeNullOrEmpty
        }

        It 'Should handle empty XML' {
            # Empty string is not accepted by the parameter, so we test with whitespace instead
            $result = ConvertFrom-DmarcXml -XmlContent ' '
            $result | Should -BeNullOrEmpty
        }

        It 'Should emit DmarcPass, Aligned_dkim, and Aligned_spf based on policy_evaluated outcome' {
            $alignmentXml = @'
<?xml version="1.0" encoding="UTF-8"?>
<feedback>
  <report_metadata>
    <org_name>test.com</org_name>
    <email>test@test.com</email>
    <report_id>alignment-test</report_id>
    <date_range><begin>1704067200</begin><end>1704153599</end></date_range>
  </report_metadata>
  <policy_published>
    <domain>example.com</domain>
    <p>none</p>
    <pct>100</pct>
  </policy_published>
  <record>
    <row>
      <source_ip>1.2.3.4</source_ip>
      <count>10</count>
      <policy_evaluated>
        <disposition>none</disposition>
        <dkim>pass</dkim>
        <spf>fail</spf>
      </policy_evaluated>
    </row>
    <identifiers><header_from>example.com</header_from></identifiers>
    <auth_results>
      <dkim><domain>example.com</domain><result>pass</result></dkim>
      <spf><domain>example.com</domain><result>fail</result></spf>
    </auth_results>
  </record>
  <record>
    <row>
      <source_ip>5.6.7.8</source_ip>
      <count>5</count>
      <policy_evaluated>
        <disposition>reject</disposition>
        <dkim>fail</dkim>
        <spf>fail</spf>
      </policy_evaluated>
    </row>
    <identifiers><header_from>example.com</header_from></identifiers>
    <auth_results>
      <dkim><domain>example.com</domain><result>fail</result></dkim>
      <spf><domain>example.com</domain><result>fail</result></spf>
    </auth_results>
  </record>
</feedback>
'@
            $result = ConvertFrom-DmarcXml -XmlContent $alignmentXml
            $result | Should -Not -BeNullOrEmpty
            $result -is [array] | Should -Be $true
            $result.Length | Should -Be 2

            # DKIM pass, SPF fail → DmarcPass = true
            $result[0].Aligned_dkim | Should -Be $true
            $result[0].Aligned_spf  | Should -Be $false
            $result[0].DmarcPass    | Should -Be $true

            # Both fail → DmarcPass = false
            $result[1].Aligned_dkim | Should -Be $false
            $result[1].Aligned_spf  | Should -Be $false
            $result[1].DmarcPass    | Should -Be $false
        }

        It 'Should derive OverrideReasonCategory from policy override reason types' {
            $overrideXml = @'
<?xml version="1.0" encoding="UTF-8"?>
<feedback>
  <report_metadata>
    <org_name>test.com</org_name>
    <email>test@test.com</email>
    <report_id>override-test</report_id>
    <date_range><begin>1704067200</begin><end>1704153599</end></date_range>
  </report_metadata>
  <policy_published>
    <domain>example.com</domain>
    <p>none</p>
    <pct>100</pct>
  </policy_published>
  <record>
    <row>
      <source_ip>10.0.0.1</source_ip>
      <count>2</count>
      <policy_evaluated>
        <disposition>none</disposition>
        <dkim>pass</dkim>
        <spf>pass</spf>
        <reason><type>mailing_list</type></reason>
      </policy_evaluated>
    </row>
    <identifiers><header_from>example.com</header_from></identifiers>
    <auth_results>
      <dkim><domain>example.com</domain><result>pass</result></dkim>
      <spf><domain>example.com</domain><result>pass</result></spf>
    </auth_results>
  </record>
  <record>
    <row>
      <source_ip>10.0.0.2</source_ip>
      <count>1</count>
      <policy_evaluated>
        <disposition>none</disposition>
        <dkim>fail</dkim>
        <spf>fail</spf>
        <reason><type>custom_receiver_logic</type></reason>
      </policy_evaluated>
    </row>
    <identifiers><header_from>example.com</header_from></identifiers>
    <auth_results>
      <dkim><domain>example.com</domain><result>fail</result></dkim>
      <spf><domain>example.com</domain><result>fail</result></spf>
    </auth_results>
  </record>
</feedback>
'@

            $result = ConvertFrom-DmarcXml -XmlContent $overrideXml
            $result[0].OverrideReasonCategory | Should -Be 'mailing_list'
            $result[1].OverrideReasonCategory | Should -Be 'other'
        }

        It 'Should emit RecordIndex starting at 0 and incrementing per record' {
            $multiXml = @'
<?xml version="1.0" encoding="UTF-8"?>
<feedback>
  <report_metadata>
    <org_name>test.com</org_name>
    <email>test@test.com</email>
    <report_id>index-test</report_id>
    <date_range><begin>1704067200</begin><end>1704153599</end></date_range>
  </report_metadata>
  <policy_published>
    <domain>example.com</domain>
    <p>none</p>
    <pct>100</pct>
  </policy_published>
  <record>
    <row>
      <source_ip>1.1.1.1</source_ip>
      <count>1</count>
      <policy_evaluated><disposition>none</disposition><dkim>pass</dkim><spf>pass</spf></policy_evaluated>
    </row>
    <identifiers><header_from>example.com</header_from></identifiers>
    <auth_results>
      <dkim><domain>example.com</domain><result>pass</result></dkim>
      <spf><domain>example.com</domain><result>pass</result></spf>
    </auth_results>
  </record>
  <record>
    <row>
      <source_ip>2.2.2.2</source_ip>
      <count>2</count>
      <policy_evaluated><disposition>none</disposition><dkim>fail</dkim><spf>fail</spf></policy_evaluated>
    </row>
    <identifiers><header_from>example.com</header_from></identifiers>
    <auth_results>
      <dkim><domain>example.com</domain><result>fail</result></dkim>
      <spf><domain>example.com</domain><result>fail</result></spf>
    </auth_results>
  </record>
</feedback>
'@
            $result = ConvertFrom-DmarcXml -XmlContent $multiXml
            $result[0].RecordIndex | Should -Be 0
            $result[1].RecordIndex | Should -Be 1
        }

        It 'Should emit a 64-character lowercase hex MessageHash per record' {
            $hashXml = @'
<?xml version="1.0" encoding="UTF-8"?>
<feedback>
  <report_metadata>
    <org_name>test.com</org_name>
    <email>test@test.com</email>
    <report_id>hash-test</report_id>
    <date_range><begin>1704067200</begin><end>1704153599</end></date_range>
  </report_metadata>
  <policy_published>
    <domain>example.com</domain>
    <p>none</p>
    <pct>100</pct>
  </policy_published>
  <record>
    <row>
      <source_ip>1.1.1.1</source_ip>
      <count>1</count>
      <policy_evaluated><disposition>none</disposition><dkim>pass</dkim><spf>pass</spf></policy_evaluated>
    </row>
    <identifiers><header_from>example.com</header_from></identifiers>
    <auth_results>
      <dkim><domain>example.com</domain><result>pass</result></dkim>
      <spf><domain>example.com</domain><result>pass</result></spf>
    </auth_results>
  </record>
</feedback>
'@
            $r1 = ConvertFrom-DmarcXml -XmlContent $hashXml -SourceMessageId 'msg-A'
            $record = if ($r1 -is [array]) { $r1[0] } else { $r1 }

            # SHA256 → 32 bytes → 64 hex chars, lowercase
            $record.MessageHash | Should -Match '^[0-9a-f]{64}$'

            # Deterministic: same inputs produce same hash
            $r2 = ConvertFrom-DmarcXml -XmlContent $hashXml -SourceMessageId 'msg-A'
            $rec2 = if ($r2 -is [array]) { $r2[0] } else { $r2 }
            $record.MessageHash | Should -Be $rec2.MessageHash

            # Different SourceMessageId → different hash
            $r3 = ConvertFrom-DmarcXml -XmlContent $hashXml -SourceMessageId 'msg-B'
            $rec3 = if ($r3 -is [array]) { $r3[0] } else { $r3 }
            $record.MessageHash | Should -Not -Be $rec3.MessageHash
        }

       It 'Should canonicalize domain and dedup fields before hashing' {
           $hashXmlA = @'
<?xml version="1.0" encoding="UTF-8"?>
<feedback>
 <report_metadata>
   <org_name> TEST.COM </org_name>
   <email>test@test.com</email>
   <report_id>Hash-Test</report_id>
   <date_range><begin>1704067200</begin><end>1704153599</end></date_range>
 </report_metadata>
 <policy_published>
   <domain>Example.Com.</domain>
   <p>none</p>
   <pct>100</pct>
 </policy_published>
 <record>
   <row>
     <source_ip> 1.1.1.1 </source_ip>
     <count>1</count>
     <policy_evaluated><disposition>none</disposition><dkim>pass</dkim><spf>pass</spf></policy_evaluated>
   </row>
   <identifiers><header_from> MAIL.EXAMPLE.COM. </header_from></identifiers>
   <auth_results>
     <dkim><domain>example.com</domain><result>pass</result></dkim>
     <spf><domain>example.com</domain><result>pass</result></spf>
   </auth_results>
 </record>
</feedback>
'@
           $hashXmlB = @'
<?xml version="1.0" encoding="UTF-8"?>
<feedback>
 <report_metadata>
   <org_name>test.com</org_name>
   <email>test@test.com</email>
   <report_id>hash-test</report_id>
   <date_range><begin>1704067200</begin><end>1704153599</end></date_range>
 </report_metadata>
 <policy_published>
   <domain>example.com</domain>
   <p>none</p>
   <pct>100</pct>
 </policy_published>
 <record>
   <row>
     <source_ip>1.1.1.1</source_ip>
     <count>1</count>
     <policy_evaluated><disposition>none</disposition><dkim>pass</dkim><spf>pass</spf></policy_evaluated>
   </row>
   <identifiers><header_from>mail.example.com</header_from></identifiers>
   <auth_results>
     <dkim><domain>example.com</domain><result>pass</result></dkim>
     <spf><domain>example.com</domain><result>pass</result></spf>
   </auth_results>
 </record>
</feedback>
'@

           $rA = ConvertFrom-DmarcXml -XmlContent $hashXmlA -SourceMessageId ' Msg-A '
           $rB = ConvertFrom-DmarcXml -XmlContent $hashXmlB -SourceMessageId 'msg-a'

           $recordA = if ($rA -is [array]) { $rA[0] } else { $rA }
           $recordB = if ($rB -is [array]) { $rB[0] } else { $rB }

           $recordA.MessageHash | Should -Be $recordB.MessageHash
           $recordA.DuplicateTelemetryKey | Should -Be $recordB.DuplicateTelemetryKey
       }

        It 'Should tolerate a non-numeric <count> value instead of failing the report' {
            $badCountXml = @'
<?xml version="1.0" encoding="UTF-8"?>
<feedback>
  <report_metadata>
    <org_name>test.com</org_name>
    <email>test@test.com</email>
    <report_id>bad-count-test</report_id>
    <date_range><begin>1704067200</begin><end>1704153599</end></date_range>
  </report_metadata>
  <policy_published>
    <domain>example.com</domain>
    <p>none</p>
    <pct>100</pct>
  </policy_published>
  <record>
    <row>
      <source_ip>1.1.1.1</source_ip>
      <count>abc</count>
      <policy_evaluated><disposition>none</disposition><dkim>pass</dkim><spf>pass</spf></policy_evaluated>
    </row>
    <identifiers><header_from>example.com</header_from></identifiers>
    <auth_results>
      <dkim><domain>example.com</domain><result>pass</result></dkim>
      <spf><domain>example.com</domain><result>pass</result></spf>
    </auth_results>
  </record>
  <record>
    <row>
      <source_ip>2.2.2.2</source_ip>
      <count>7</count>
      <policy_evaluated><disposition>none</disposition><dkim>pass</dkim><spf>pass</spf></policy_evaluated>
    </row>
    <identifiers><header_from>example.com</header_from></identifiers>
    <auth_results>
      <dkim><domain>example.com</domain><result>pass</result></dkim>
      <spf><domain>example.com</domain><result>pass</result></spf>
    </auth_results>
  </record>
</feedback>
'@
            $result = @(ConvertFrom-DmarcXml -XmlContent $badCountXml -WarningAction SilentlyContinue)
            $result.Length | Should -Be 2
            $result[0].MessageCount | Should -Be 0
            $result[1].MessageCount | Should -Be 7
        }

        It 'Should default a non-numeric <pct> value to 100 instead of failing the report' {
            $badPctXml = @'
<?xml version="1.0" encoding="UTF-8"?>
<feedback>
  <report_metadata>
    <org_name>test.com</org_name>
    <email>test@test.com</email>
    <report_id>bad-pct-test</report_id>
    <date_range><begin>1704067200</begin><end>1704153599</end></date_range>
  </report_metadata>
  <policy_published>
    <domain>example.com</domain>
    <p>none</p>
    <pct>often</pct>
  </policy_published>
  <record>
    <row>
      <source_ip>1.1.1.1</source_ip>
      <count>1</count>
      <policy_evaluated><disposition>none</disposition><dkim>pass</dkim><spf>pass</spf></policy_evaluated>
    </row>
    <identifiers><header_from>example.com</header_from></identifiers>
    <auth_results>
      <dkim><domain>example.com</domain><result>pass</result></dkim>
      <spf><domain>example.com</domain><result>pass</result></spf>
    </auth_results>
  </record>
</feedback>
'@
            $result = ConvertFrom-DmarcXml -XmlContent $badPctXml -WarningAction SilentlyContinue
            $record = if ($result -is [array]) { $result[0] } else { $result }
            $record.PolicyPublished_pct | Should -Be 100
        }

        It 'Should prohibit DTD processing (security check)' {
            $dtdXml = @'
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE feedback [<!ENTITY xxe SYSTEM "file:///etc/passwd">]>
<feedback>
  <report_metadata>
    <org_name>&xxe;</org_name>
  </report_metadata>
</feedback>
'@
            $result = ConvertFrom-DmarcXml -XmlContent $dtdXml
            # Should either return empty or fail safely (no file read)
            $result | Should -BeNullOrEmpty
        }
    }

    Context 'Expand-DmarcAttachments' {
        It 'Should handle XML attachments and return Xml key' {
            $xmlContent = @'
<?xml version="1.0" encoding="UTF-8"?>
<feedback>
  <report_metadata>
    <org_name>test</org_name>
    <report_id>123</report_id>
  </report_metadata>
  <policy_published>
    <domain>example.com</domain>
    <p>none</p>
  </policy_published>
</feedback>
'@
            $xmlBytes = [System.Text.Encoding]::UTF8.GetBytes($xmlContent)
            $base64 = [System.Convert]::ToBase64String($xmlBytes)

            $attachment = @{
                '@odata.type' = '#microsoft.graph.fileAttachment'
                'name' = 'report.xml'
                'contentBytes' = $base64
            }

            $result = Expand-DmarcAttachments -Attachments @($attachment)
            $result | Should -Not -BeNullOrEmpty
            $result.Xml | Should -Not -BeNullOrEmpty
            $result.Xml.Count | Should -Be 1
            $result.Xml[0] | Should -BeLike '*<feedback>*'
        }

        It 'Should handle GZIP attachments' {
            $xmlContent = '<feedback><report_metadata><org_name>test</org_name></report_metadata></feedback>'
            $xmlBytes = [System.Text.Encoding]::UTF8.GetBytes($xmlContent)

            # Compress to GZIP
            $memStream = [System.IO.MemoryStream]::new()
            $gzipStream = [System.IO.Compression.GZipStream]::new($memStream, [System.IO.Compression.CompressionMode]::Compress)
            $gzipStream.Write($xmlBytes, 0, $xmlBytes.Length)
            $gzipStream.Close()
            $gzipBytes = $memStream.ToArray()
            $memStream.Dispose()

            $base64 = [System.Convert]::ToBase64String($gzipBytes)
            $attachment = @{
                '@odata.type' = '#microsoft.graph.fileAttachment'
                'name' = 'report.xml.gz'
                'contentBytes' = $base64
            }

            $result = Expand-DmarcAttachments -Attachments @($attachment)
            $result | Should -Not -BeNullOrEmpty
            $result.Xml | Should -Not -BeNullOrEmpty
            $result.Xml.Count | Should -Be 1
            $result.Xml[0] | Should -BeLike '*<feedback>*'
        }

        It 'Should skip oversized attachments' {
            # Create a very large content exceeding MaxAttachmentBytes
            $largeBytes = [byte[]]::new(26 * 1024 * 1024)  # 26 MB
            $base64 = [System.Convert]::ToBase64String($largeBytes)

            $attachment = @{
                '@odata.type' = '#microsoft.graph.fileAttachment'
                'name' = 'large.xml'
                'contentBytes' = $base64
            }

            $result = Expand-DmarcAttachments -Attachments @($attachment)
            $result.Xml.Count | Should -Be 0
        }

        It 'Should skip non-file attachments' {
            $attachment = @{
                '@odata.type' = '#microsoft.graph.itemAttachment'
                'name' = 'meeting.ics'
            }

            $result = Expand-DmarcAttachments -Attachments @($attachment)
            $result.Xml.Count | Should -Be 0
        }

        It 'Should skip unrecognized file extensions' {
            $content = 'test content'
            $contentBytes = [System.Text.Encoding]::UTF8.GetBytes($content)
            $base64 = [System.Convert]::ToBase64String($contentBytes)

            $attachment = @{
                '@odata.type' = '#microsoft.graph.fileAttachment'
                'name' = 'document.pdf'
                'contentBytes' = $base64
            }

            $result = Expand-DmarcAttachments -Attachments @($attachment)
            $result.Xml.Count | Should -Be 0
        }
    }

    Context 'Cumulative decompression limit' {
        BeforeAll {
            # Shrink the message-level budget so the cap can be exercised with small fixtures.
            InModuleScope DmarcHelpers {
                $script:SavedMaxTotalDecompressedBytes = $script:MaxTotalDecompressedBytes
                $script:MaxTotalDecompressedBytes = 1024
            }
        }

        AfterAll {
            InModuleScope DmarcHelpers {
                $script:MaxTotalDecompressedBytes = $script:SavedMaxTotalDecompressedBytes
                Remove-Variable -Name SavedMaxTotalDecompressedBytes -Scope Script
            }
        }

        It 'Should stop adding content once the total decompressed limit is exceeded' {
            $xml = '<feedback>' + ('x' * 900) + '</feedback>'
            $base64 = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($xml))

            $attachments = @(
                @{ '@odata.type' = '#microsoft.graph.fileAttachment'; 'name' = 'r1.xml'; 'contentBytes' = $base64 }
                @{ '@odata.type' = '#microsoft.graph.fileAttachment'; 'name' = 'r2.xml'; 'contentBytes' = $base64 }
            )

            $result = Expand-DmarcAttachments -Attachments $attachments -WarningAction SilentlyContinue
            # First attachment fits the 1024-byte budget; the second exceeds it.
            $result.Xml.Count | Should -Be 1
        }

        It 'Should enforce the remaining budget while decompressing standalone GZIP attachments' {
            $xml = '<feedback>' + ('x' * 2000) + '</feedback>'
            $xmlBytes = [System.Text.Encoding]::UTF8.GetBytes($xml)

            $memStream = [System.IO.MemoryStream]::new()
            $gzipStream = [System.IO.Compression.GZipStream]::new($memStream, [System.IO.Compression.CompressionMode]::Compress)
            $gzipStream.Write($xmlBytes, 0, $xmlBytes.Length)
            $gzipStream.Close()
            $gzipBytes = $memStream.ToArray()
            $memStream.Dispose()

            $attachment = @{
                '@odata.type' = '#microsoft.graph.fileAttachment'
                'name' = 'report.xml.gz'
                'contentBytes' = [System.Convert]::ToBase64String($gzipBytes)
            }

            # Decompressed size (2021 bytes) exceeds the 1024-byte budget, so the
            # copy must abort during extraction rather than after.
            $result = Expand-DmarcAttachments -Attachments @($attachment) -WarningAction SilentlyContinue
            $result.Xml.Count | Should -Be 0
        }

        It 'Should enforce the budget across entries within a single ZIP archive' {
            $xml = '<feedback>' + ('x' * 600) + '</feedback>'
            $xmlBytes = [System.Text.Encoding]::UTF8.GetBytes($xml)

            $zipStream = [System.IO.MemoryStream]::new()
            $archive = [System.IO.Compression.ZipArchive]::new($zipStream, [System.IO.Compression.ZipArchiveMode]::Create, $true)
            foreach ($entryName in @('a.xml', 'b.xml', 'c.xml')) {
                $entry = $archive.CreateEntry($entryName)
                $entryStream = $entry.Open()
                $entryStream.Write($xmlBytes, 0, $xmlBytes.Length)
                $entryStream.Dispose()
            }
            $archive.Dispose()
            $zipBytes = $zipStream.ToArray()
            $zipStream.Dispose()

            $attachment = @{
                '@odata.type' = '#microsoft.graph.fileAttachment'
                'name' = 'reports.zip'
                'contentBytes' = [System.Convert]::ToBase64String($zipBytes)
            }

            $result = Expand-DmarcAttachments -Attachments @($attachment) -WarningAction SilentlyContinue
            # Only the first 600-byte entry fits within the 1024-byte budget.
            $result.Xml.Count | Should -Be 1
        }
    }

    Context 'ConvertTo-SafeLogText' {
        It 'Should strip control characters including CR/LF' {
            $sanitized = ConvertTo-SafeLogText -Text "line1`r`nFAKE LOG ENTRY`tend"
            $sanitized | Should -Not -Match '[\r\n\t]'
            $sanitized | Should -BeLike '*line1*FAKE LOG ENTRY*end*'
        }

        It 'Should truncate long values' {
            $sanitized = ConvertTo-SafeLogText -Text ('a' * 500) -MaxLength 100
            $sanitized.Length | Should -BeLessOrEqual 103
        }

        It 'Should return an empty string for null or empty input' {
            ConvertTo-SafeLogText -Text $null | Should -Be ''
            ConvertTo-SafeLogText -Text '' | Should -Be ''
        }
    }

    Context 'Send-DmarcRecordsToLogAnalytics' {
        It 'Should require DCR_ENDPOINT environment variable' {
            $env:DCR_ENDPOINT = $null
            $env:DCR_IMMUTABLE_ID = 'test'
            $env:DCR_STREAM_NAME = 'test'

            $records = @(@{ TestField = 'value' })
            { Send-DmarcRecordsToLogAnalytics -Records $records } | Should -Throw '*DCR*'
        }

        It 'Should require DCR_IMMUTABLE_ID environment variable' {
            $env:DCR_ENDPOINT = 'https://test.endpoint'
            $env:DCR_IMMUTABLE_ID = $null
            $env:DCR_STREAM_NAME = 'test'

            $records = @(@{ TestField = 'value' })
            { Send-DmarcRecordsToLogAnalytics -Records $records } | Should -Throw '*DCR*'
        }

        It 'Should require DCR_STREAM_NAME environment variable' {
            $env:DCR_ENDPOINT = 'https://test.endpoint'
            $env:DCR_IMMUTABLE_ID = 'test-id'
            $env:DCR_STREAM_NAME = $null

            $records = @(@{ TestField = 'value' })
            { Send-DmarcRecordsToLogAnalytics -Records $records } | Should -Throw '*DCR*'
        }
    }

    Context 'Invoke-WithRetry (via Invoke-GraphRequest)' {
        It 'Should succeed on the first attempt when no error occurs' {
            $env:IDENTITY_ENDPOINT = 'https://identity.endpoint'
            $env:IDENTITY_HEADER   = 'test-header'

            $callCount = 0
            Mock -ModuleName DmarcHelpers Invoke-RestMethod {
                $callCount++
                if ($callCount -eq 1 -and $Uri -like '*identity*') { return @{ access_token = 'tok' } }
                return @{ value = 'ok' }
            }

            $result = Invoke-GraphRequest -Uri 'https://graph.microsoft.com/v1.0/me' -Token 'test-token'
            $result.value | Should -Be 'ok'
        }

        It 'Should rethrow immediately on non-retryable errors (e.g. 404)' {
            $env:IDENTITY_ENDPOINT = 'https://identity.endpoint'
            $env:IDENTITY_HEADER   = 'test-header'

            $script:attempts = 0
            Mock -ModuleName DmarcHelpers Invoke-RestMethod {
                $script:attempts++
                $response = [System.Net.Http.HttpResponseMessage]::new([System.Net.HttpStatusCode]::NotFound)
                throw [Microsoft.PowerShell.Commands.HttpResponseException]::new('404', $response)
            }
            Mock -ModuleName DmarcHelpers Start-Sleep { }

            { Invoke-GraphRequest -Uri 'https://graph.microsoft.com/v1.0/me' -Token 'test-token' } | Should -Throw
            $script:attempts | Should -Be 1
        }

        It 'Should retry on 429 up to MaxAttempts then rethrow' {
            $env:IDENTITY_ENDPOINT = 'https://identity.endpoint'
            $env:IDENTITY_HEADER   = 'test-header'

            $script:attempts = 0
            Mock -ModuleName DmarcHelpers Invoke-RestMethod {
                $script:attempts++
                $response = [System.Net.Http.HttpResponseMessage]::new([System.Net.HttpStatusCode]::TooManyRequests)
                throw [Microsoft.PowerShell.Commands.HttpResponseException]::new('429', $response)
            }
            Mock -ModuleName DmarcHelpers Start-Sleep { }

            { Invoke-GraphRequest -Uri 'https://graph.microsoft.com/v1.0/me' -Token 'test-token' } | Should -Throw
            # Default MaxAttempts = 4
            $script:attempts | Should -Be 4
        }

        It 'Should succeed after a transient 503 on the first attempt' {
            $env:IDENTITY_ENDPOINT = 'https://identity.endpoint'
            $env:IDENTITY_HEADER   = 'test-header'

            $script:attempts = 0
            Mock -ModuleName DmarcHelpers Invoke-RestMethod {
                $script:attempts++
                if ($script:attempts -eq 1) {
                    $response = [System.Net.Http.HttpResponseMessage]::new([System.Net.HttpStatusCode]::ServiceUnavailable)
                    throw [Microsoft.PowerShell.Commands.HttpResponseException]::new('503', $response)
                }
                return @{ value = 'recovered' }
            }
            Mock -ModuleName DmarcHelpers Start-Sleep { }

            $result = Invoke-GraphRequest -Uri 'https://graph.microsoft.com/v1.0/me' -Token 'test-token'
            $result.value | Should -Be 'recovered'
            $script:attempts | Should -Be 2
        }
    }

}
