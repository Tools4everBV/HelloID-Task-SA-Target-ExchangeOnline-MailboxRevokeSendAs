# HelloID-Task-SA-Target-ExchangeOnline-MailboxRevokeSendAs
###########################################################
# Form mapping
$formObject = @{
    MailboxIdentity = $form.MailboxIdentity
    UsersToRemove   = $form.UsersToRemove.id
}
[bool]$IsConnected = $false

try {
    Write-Information "Executing ExchangeOnline action: [MailboxRevokeSendAs] for: [$($formObject.MailboxIdentity)]"

    $null = Import-Module ExchangeOnlineManagement

    $securePassword = ConvertTo-SecureString $ExchangeOnlineAdminPassword -AsPlainText -Force
    $credential = [System.Management.Automation.PSCredential]::new($ExchangeOnlineAdminUsername, $securePassword)
    $null = Connect-ExchangeOnline -Credential $credential -ShowBanner:$false -ShowProgress:$false -ErrorAction Stop -Verbose:$false -CommandName 'Remove-RecipientPermission', 'Disconnect-ExchangeOnline'
    $IsConnected = $true

    foreach ($user in $formObject.UsersToRemove) {
        $null =  Remove-RecipientPermission -Identity $formObject.MailboxIdentity -AccessRights SendAs -Confirm:$false -Trustee $user -ErrorAction Stop

        $auditLog = @{
            Action            = 'UpdateResource'
            System            = 'ExchangeOnline'
            TargetIdentifier  = $formObject.MailboxIdentity
            TargetDisplayName = $formObject.MailboxIdentity
            Message           = "ExchangeOnline action: [MailboxRevokeSendAs] Revoke [$($user)] from mailbox [$($formObject.MailboxIdentity)] executed successfully"
            IsError           = $false
        }
        Write-Information -Tags 'Audit' -MessageData $auditLog
        Write-Information "ExchangeOnline action: [MailboxRevokeSendAs] Revoke [$($user)] from mailbox [$($formObject.MailboxIdentity)] executed successfully"
    }
} catch {
    $ex = $_
    $auditLog = @{
        Action            = 'UpdateResource'
        System            = 'ExchangeOnline'
        TargetIdentifier  = $formObject.MailboxIdentity
        TargetDisplayName = $formObject.MailboxIdentity
        Message           = "Could not execute ExchangeOnline action: [MailboxRevokeSendAs] for: [$($formObject.MailboxIdentity)], error: $($ex.Exception.Message)"
        IsError           = $true
    }
    Write-Information -Tags "Audit" -MessageData $auditLog
    Write-Error "Could not execute ExchangeOnline action: [MailboxRevokeSendAs] for: [$($formObject.MailboxIdentity)], error: $($ex.Exception.Message)"
} finally {
    if ($IsConnected) {
        $null = Disconnect-ExchangeOnline -Confirm:$false -Verbose:$false
    }
}
###########################################################
