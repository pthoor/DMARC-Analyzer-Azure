#Requires -Version 7.4
#Requires -Modules Pester

<#
.SYNOPSIS
    Pester tests for New-GraphSubscription.ps1 script.
.DESCRIPTION
    Tests the Graph subscription setup script for proper parameter validation,
    error handling, and workflow.
#>

BeforeAll {
    $scriptPath = "$PSScriptRoot/../scripts/New-GraphSubscription.ps1"
}

Describe 'New-GraphSubscription.ps1' {
    Context 'Script Structure' {
        It 'Should exist' {
            Test-Path $scriptPath | Should -Be $true
        }

        It 'Should have valid PowerShell syntax' {
            $errors = $null
            $null = [System.Management.Automation.PSParser]::Tokenize(
                (Get-Content $scriptPath -Raw), 
                [ref]$errors
            )
            $errors | Should -BeNullOrEmpty
        }

        It 'Should have CmdletBinding attribute' {
            $content = Get-Content $scriptPath -Raw
            $content | Should -Match '\[CmdletBinding\(\)\]'
        }

        It 'Should have required parameters' {
            $content = Get-Content $scriptPath -Raw
            $content | Should -Match '(?s)\[Parameter\(Mandatory\)\].*?\$FunctionAppName'
            $content | Should -Match '(?s)\[Parameter\(Mandatory\)\].*?\$ResourceGroupName'
            $content | Should -Match '(?s)\[Parameter\(Mandatory\)\].*?\$MailboxUserId'
            $content | Should -Match '(?s)\[Parameter\(Mandatory\)\].*?\$SubscriptionId'
            $content | Should -Match '(?s)\[Parameter\(Mandatory\)\].*?\$GraphClientState'
        }

        It 'Should set ErrorActionPreference to Stop' {
            $content = Get-Content $scriptPath -Raw
            $content | Should -Match '\$ErrorActionPreference\s*=\s*[''"]Stop[''"]'
        }
    }

    Context 'Parameter Validation' {
        It 'Should validate resource group early in prerequisites' {
            $content = Get-Content $scriptPath -Raw
            # Check that Get-AzResourceGroup is called in the prerequisites section
            $content | Should -Match 'Get-AzResourceGroup.*-Name.*\$ResourceGroupName'
        }

        It 'Should check prerequisite completion before continuing' {
            $content = Get-Content $scriptPath -Raw
            # Check for user confirmation prompt
            $content | Should -Match 'Read-Host.*completed'
        }

        It 'Should construct Event Grid notification URL correctly' {
            $content = Get-Content $scriptPath -Raw
            $content | Should -Match 'EventGrid:\?azuresubscriptionid='
            $content | Should -Match 'resourcegroup='
            $content | Should -Match 'partnertopic='
            $content | Should -Match 'location='
        }
    }

    Context 'Security Considerations' {
        It 'Should handle GraphClientState securely' {
            $content = Get-Content $scriptPath -Raw
            # Verify that clientState is used in the subscription body
            $content | Should -Match 'clientState.*=.*\$GraphClientState'
        }

        It 'Should set appropriate subscription expiration' {
            $content = Get-Content $scriptPath -Raw
            # Check that expiration is set (4200 minutes is used, which is under the 4230 max)
            $content | Should -Match 'expirationDateTime'
            $content | Should -Match 'AddMinutes\(4200\)'
        }
    }

    Context 'Graph API Integration' {
        It 'Should use Microsoft Graph API v1.0' {
            $content = Get-Content $scriptPath -Raw
            $content | Should -Match 'https://graph\.microsoft\.com/v1\.0'
        }

        It 'Should create subscription with correct properties' {
            $content = Get-Content $scriptPath -Raw
            $content | Should -Match 'changeType'
            $content | Should -Match 'notificationUrl'
            $content | Should -Match 'resource'
            $content | Should -Match 'expirationDateTime'
            $content | Should -Match 'clientState'
        }

        It 'Should save subscription ID to Function App settings' {
            $content = Get-Content $scriptPath -Raw
            $content | Should -Match 'GRAPH_SUBSCRIPTION_ID'
            $content | Should -Match 'Set-AzWebApp'
        }

        It 'Should verify that subscription ID was saved' {
            $content = Get-Content $scriptPath -Raw
            # Check for verification after saving
            $content | Should -Match 'Get-AzWebApp.*-Name.*\$FunctionAppName'
            $content | Should -Match 'GRAPH_SUBSCRIPTION_ID.*-eq.*\$graphSubscriptionId'
        }
    }

    Context 'User Experience' {
        It 'Should provide clear output messages' {
            $content = Get-Content $scriptPath -Raw
            $content | Should -Match 'Write-Host'
        }

        It 'Should provide next steps guidance' {
            $content = Get-Content $scriptPath -Raw
            $content | Should -Match 'Remaining manual steps'
            $content | Should -Match 'Activate.*partner topic'
        }

        It 'Should display created subscription ID' {
            $content = Get-Content $scriptPath -Raw
            $content | Should -Match 'Graph subscription ID:'
        }
    }
}
