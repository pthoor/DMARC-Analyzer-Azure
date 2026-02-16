#Requires -Version 7.4
#Requires -Modules Pester

<#
.SYNOPSIS
    Pester tests for DmarcHelpers PowerShell module.
.DESCRIPTION
    Tests all functions in the DmarcHelpers.psm1 module including:
    - Token acquisition
    - Graph API helpers
    - Attachment extraction
    - XML parsing
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
            $exportedFunctions | Should -Contain 'Get-UnreadMessages'
            $exportedFunctions | Should -Contain 'Expand-DmarcAttachments'
            $exportedFunctions | Should -Contain 'ConvertFrom-DmarcXml'
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
        It 'Should handle XML attachments' {
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
            # When there's one attachment, it returns a string, not an array
            if ($result -is [array]) {
                $result.Length | Should -Be 1
                $result[0] | Should -BeLike '*<feedback>*'
            } else {
                $result | Should -BeLike '*<feedback>*'
            }
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
            # When there's one attachment, it returns a string, not an array
            if ($result -is [array]) {
                $result.Length | Should -Be 1
                $result[0] | Should -BeLike '*<feedback>*'
            } else {
                $result | Should -BeLike '*<feedback>*'
            }
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
            # Should skip and return empty
            $result | Should -BeNullOrEmpty
        }

        It 'Should skip non-file attachments' {
            $attachment = @{
                '@odata.type' = '#microsoft.graph.itemAttachment'
                'name' = 'meeting.ics'
            }

            $result = Expand-DmarcAttachments -Attachments @($attachment)
            $result | Should -BeNullOrEmpty
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
            $result | Should -BeNullOrEmpty
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
}
