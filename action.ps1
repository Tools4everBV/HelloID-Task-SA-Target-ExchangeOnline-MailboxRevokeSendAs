# HelloID-Task-SA-Target-ExchangeOnline-MailboxRevokeSendAs
###########################################################
# Form mapping
$formObject = @{
    MailboxDistinguishedName = $form.MailboxDistinguishedName
    UsersToRemove            = $form.UsersToRemove.id
}
[bool]$IsConnected = $false

try {
    Write-Information "Executing ExchangeOnline action: [MailboxRevokeSendAs] for: [$($formObject.MailboxDistinguishedName)]"

    $null = Import-Module ExchangeOnlineManagement

    $securePassword = ConvertTo-SecureString $ExchangeOnlineAdminPassword -AsPlainText -Force
    $credential = [System.Management.Automation.PSCredential]::new($ExchangeOnlineAdminUsername, $securePassword)
    $null = Connect-ExchangeOnline -Credential $credential -ShowBanner:$false -ShowProgress:$false -ErrorAction Stop -Verbose:$false -CommandName 'Remove-RecipientPermission', 'Disconnect-ExchangeOnline'
    $IsConnected = $true

    foreach ($user in $formObject.UsersToRemove) {
        $null =  Remove-RecipientPermission -Identity $formObject.MailboxDistinguishedName -AccessRights SendAs -Confirm:$false -Trustee $user -ErrorAction Stop

        $auditLog = @{
            Action            = 'UpdateResource'
            System            = 'ExchangeOnline'
            TargetIdentifier  = $formObject.MailboxDistinguishedName
            TargetDisplayName = $formObject.MailboxDistinguishedName
            Message           = "ExchangeOnline action: [MailboxRevokeSendAs] Revoke [$($user)] from mailbox [$($formObject.MailboxDistinguishedName)] executed successfully"
            IsError           = $false
        }
        Write-Information -Tags 'Audit' -MessageData $auditLog
        Write-Information "ExchangeOnline action: [MailboxRevokeSendAs] Revoke [$($user)] from mailbox [$($formObject.MailboxDistinguishedName)] executed successfully"
    }
} catch {
    $ex = $_
    $auditLog = @{
        Action            = 'UpdateResource'
        System            = 'ExchangeOnline'
        TargetIdentifier  = $formObject.MailboxDistinguishedName
        TargetDisplayName = $formObject.MailboxDistinguishedName
        Message           = "Could not execute ExchangeOnline action: [MailboxRevokeSendAs] for: [$($formObject.MailboxDistinguishedName)], error: $($ex.Exception.Message)"
        IsError           = $true
    }
    Write-Information -Tags "Audit" -MessageData $auditLog
    Write-Error "Could not execute ExchangeOnline action: [MailboxRevokeSendAs] for: [$($formObject.MailboxDistinguishedName)], error: $($ex.Exception.Message)"
} finally {
    if ($IsConnected) {
        $null = Disconnect-ExchangeOnline -Confirm:$false -Verbose:$false
    }
}
###########################################################
