#HelloID-Task-SA-Target-ExchangeOnPremises-SharedMailboxGrantFullAccess
######################################################################

# Form mapping
$formObject = @{
    Users    = $form.Users
    MailboxIdentity = $form.MailboxIdentity
}

[bool]$IsConnected = $false
try {
    $adminSecurePassword = ConvertTo-SecureString -String $ExchangeAdminPassword -AsPlainText -Force
    $adminCredential = [System.Management.Automation.PSCredential]::new($ExhangeAdminUsername, $adminSecurePassword)
    $sessionOption = New-PSSessionOption -SkipCACheck -SkipCNCheck
    $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ExchangeConnectionUri -Credential $adminCredential -SessionOption $sessionOption -Authentication Kerberos  -ErrorAction Stop
    Import-PSSession $exchangeSession -DisableNameChecking -AllowClobber -CommandName 'Add-MailboxPermission'

    $IsConnected = $true

    foreach ($user in $formObject.Users) {
        Write-Information "Executing ExchangeOnPremises action: [SharedMailboxGrantFullAccess] for: [$user]"

        $ParamsAddMailboxPermission = @{
            Identity        = $formObject.MailboxIdentity
            User            = $user
            AccessRights    = "FullAccess"
            InheritanceType = "All"
        }
        
        $null = Add-MailboxPermission @ParamsAddMailboxPermission

        $auditLog = @{
            Action            = 'UpdateResource'
            System            = 'ExchangeOnPremises'
            TargetIdentifier  = $formObject.MailboxIdentity
            TargetDisplayName = $formObject.MailboxIdentity
            Message           = "ExchangeOnPremises action: [SharedMailboxGrantFullAccess][$($user)] to mailbox [$($formObject.MailboxIdentity)] executed successfully"
            IsError           = $false
        }
        Write-Information -Tags 'Audit' -MessageData $auditLog
        Write-Information "ExchangeOnPremises action: [SharedMailboxGrantFullAccess][$($user)] to mailbox [$($formObject.MailboxIdentity)] executed successfully"
    }
}
catch {
    $ex = $_
    $auditLog = @{
        Action            = 'UpdateResource'
        System            = 'ExchangeOnPremises'
        TargetIdentifier  = $formObject.MailboxIdentity
        TargetDisplayName = $formObject.MailboxIdentity
        Message           = "Could not execute Exchange action: [SharedMailboxGrantFullAccess] for: [$($formObject.MailboxIdentity)], error: $($ex.Exception.Message)"
        IsError           = $true
    }
    Write-Information -Tags 'Audit' -MessageData $auditLog
    Write-Error "Could not execute Exchange action: [SharedMailboxGrantFullAccess] for: [$($formObject.MailboxIdentity)], error: $($ex.Exception.Message)"
}
finally {
    if ($IsConnected) {
        Remove-PsSession -Session $exchangeSession -Confirm:$false  -ErrorAction Stop
    }
}