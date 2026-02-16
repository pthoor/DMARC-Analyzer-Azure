#Requires -Version 7.4
#Requires -Modules Pester

<#
.SYNOPSIS
    Pester tests for Grant-MIExchangeRBAC.ps1 script.
.DESCRIPTION
    Tests the Exchange RBAC setup script for proper parameter validation,
    security controls, and workflow.
#>

BeforeAll {
    $scriptPath = "$PSScriptRoot/../scripts/Grant-MIExchangeRBAC.ps1"
}

Describe 'Grant-MIExchangeRBAC.ps1' {
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
            $content | Should -Match '(?s)\[Parameter\(Mandatory\)\].*?\$MailboxAddress'
        }

        It 'Should set ErrorActionPreference to Stop' {
            $content = Get-Content $scriptPath -Raw
            $content | Should -Match '\$ErrorActionPreference\s*=\s*[''"]Stop[''"]'
        }
    }

    Context 'Managed Identity Operations' {
        It 'Should retrieve Function App Managed Identity' {
            $content = Get-Content $scriptPath -Raw
            $content | Should -Match 'Get-AzWebApp.*-Name.*\$FunctionAppName'
            $content | Should -Match '\.Identity\.PrincipalId'
        }

        It 'Should validate MI is enabled' {
            $content = Get-Content $scriptPath -Raw
            $content | Should -Match 'if.*-not.*\$miPrincipalId'
            $content | Should -Match 'Managed Identity'
        }

        It 'Should get service principal details' {
            $content = Get-Content $scriptPath -Raw
            $content | Should -Match 'Get-AzADServicePrincipal.*-ObjectId'
        }
    }

    Context 'Microsoft Graph API Roles' {
        It 'Should assign Mail.Read and Mail.ReadWrite roles' {
            $content = Get-Content $scriptPath -Raw
            $content | Should -Match 'Mail\.Read'
            $content | Should -Match 'Mail\.ReadWrite'
        }

        It 'Should check for existing role assignments' {
            $content = Get-Content $scriptPath -Raw
            $content | Should -Match 'Get-MgServicePrincipalAppRoleAssignment'
        }

        It 'Should create new role assignments if needed' {
            $content = Get-Content $scriptPath -Raw
            $content | Should -Match 'New-MgServicePrincipalAppRoleAssignment'
        }

        It 'Should connect to Microsoft Graph if not connected' {
            $content = Get-Content $scriptPath -Raw
            $content | Should -Match 'Get-MgContext'
            $content | Should -Match 'Connect-MgGraph'
        }
    }

    Context 'Exchange Online RBAC' {
        It 'Should create Exchange service principal' {
            $content = Get-Content $scriptPath -Raw
            $content | Should -Match 'New-ServicePrincipal'
            $content | Should -Match '-AppId.*\$miAppId'
        }

        It 'Should check if service principal already exists' {
            $content = Get-Content $scriptPath -Raw
            $content | Should -Match 'Get-ServicePrincipal.*-Identity.*\$miAppId'
        }

        It 'Should create Management Scope for mailbox restriction' {
            $content = Get-Content $scriptPath -Raw
            $content | Should -Match 'New-ManagementScope'
            $content | Should -Match 'RecipientRestrictionFilter'
            $content | Should -Match 'PrimarySmtpAddress -eq'
        }

        It 'Should validate mailbox email address format' {
            $content = Get-Content $scriptPath -Raw
            $content | Should -Match '\[System\.Net\.Mail\.MailAddress\]'
        }

        It 'Should assign Application RBAC roles' {
            $content = Get-Content $scriptPath -Raw
            $content | Should -Match 'Application Mail\.Read'
            $content | Should -Match 'Application Mail\.ReadWrite'
            $content | Should -Match 'New-ManagementRoleAssignment'
        }

        It 'Should scope RBAC roles to the specific mailbox' {
            $content = Get-Content $scriptPath -Raw
            $content | Should -Match 'CustomResourceScope.*\$scopeName'
        }

        It 'Should connect to Exchange Online if not connected' {
            $content = Get-Content $scriptPath -Raw
            $content | Should -Match 'Get-OrganizationConfig'
            $content | Should -Match 'Connect-ExchangeOnline'
        }
    }

    Context 'Security Considerations' {
        It 'Should implement least privilege with Management Scope' {
            $content = Get-Content $scriptPath -Raw
            # Verify that scope restricts to specific mailbox
            $content | Should -Match 'PrimarySmtpAddress -eq.*\$MailboxAddress'
        }

        It 'Should provide permission propagation notice' {
            $content = Get-Content $scriptPath -Raw
            $content | Should -Match '30 min.*2 hours.*propagate'
        }

        It 'Should provide verification command' {
            $content = Get-Content $scriptPath -Raw
            $content | Should -Match 'Test-ServicePrincipalAuthorization'
        }
    }

    Context 'Error Handling' {
        It 'Should handle missing Managed Identity' {
            $content = Get-Content $scriptPath -Raw
            $content | Should -Match 'throw.*Managed Identity'
        }

        It 'Should handle invalid email address' {
            $content = Get-Content $scriptPath -Raw
            $content | Should -Match 'throw.*email address'
        }

        It 'Should check for existing assignments and handle gracefully' {
            $content = Get-Content $scriptPath -Raw
            $content | Should -Match 'Get-ManagementRoleAssignment'
            $content | Should -Match '-ErrorAction SilentlyContinue'
        }
    }

    Context 'User Experience' {
        It 'Should provide clear progress indicators' {
            $content = Get-Content $scriptPath -Raw
            $content | Should -Match '\[1/5\].*Retrieving Managed Identity'
            $content | Should -Match '\[2/5\].*Graph app roles'
            $content | Should -Match '\[3/5\].*service principal'
            $content | Should -Match '\[4/5\].*Management Scope'
            $content | Should -Match '\[5/5\].*RBAC roles'
        }

        It 'Should provide summary output' {
            $content = Get-Content $scriptPath -Raw
            $content | Should -Match 'Setup complete'
            $content | Should -Match 'Graph API:.*Mail\.Read'
            $content | Should -Match 'Exchange RBAC:.*Scoped'
        }
    }
}
