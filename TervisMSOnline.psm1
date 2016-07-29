Function Get-TempPassword() {
    Param(
        [int]$length=120,
        [string[]]$sourcedata
    )

    For ($loop=1; $loop –le $length; $loop++) {
        $TempPassword+=($sourcedata | Get-Random)
    }
    Return $TempPassword
}

function Remove-TervisMSOLUser{
    param(
        [parameter(mandatory)]$Identity,
        [parameter(mandatory)]$DirSyncServer
    )
    <# 
    You must install the "Microsoft Online Services Sign-In Assistant for IT Professionals RTW" and 
    the "Azure Active Directory Module for Windows PowerShell (64-bit version)" before you can run this.
    The links are below.
    http://go.microsoft.com/fwlink/?LinkID=286152
    http://go.microsoft.com/fwlink/p/?linkid=236297
    You must also set the $DirSyncServer variable to the DirSync server in your environment 
    #>

    $UserObject = get-aduser $Identity -properties DistinguishedName,UserPrincipalName,Manager
    $DN = $UserObject | select -ExpandProperty DistinguishedName
    $UserPrincipalName = $UserObject | select -ExpandProperty UserPrincipalName
    $ManagerDn = $UserObject | select -ExpandProperty Manager
    $ManagerUpn = get-aduser $ManagerDn | select -ExpandProperty UserPrincipalName

    # Connect to Exchange Online with your user@domain.com credentials
    write-verbose "Connect to Exchange Online with your user@domain.com credentials"
    $Credential = Import-Clixml $env:USERPROFILE\ExchangeOnlineCredential.txt
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -Authentication Basic -ConnectionUri https://ps.outlook.com/powershell -AllowRedirection:$true -Credential $credential
    Import-PSSession $Session -DisableNameChecking

    # Remove Active Sync Devices
    Write-Verbose "Removing Users Active Sync Devices"
    if (Get-ActiveSyncDevice -Mailbox $UserPrincipalName -ErrorAction SilentlyContinue) {
        Get-ActiveSyncDevice -Mailbox $UserPrincipalName | Remove-ActiveSyncDevice
    }

    # Convert the users mailbox to a shared mailbox
    Write-Verbose "Convert the users mailbox to a shared mailbox"
    Set-Mailbox $UserPrincipalName -Type shared

    # Grant shared mailbox permissions to their manager
    if ($ManagerUpn) {
        Add-MailboxPermission -Identity $UserPrincipalName -User $ManagerUpn -AccessRights FullAccess -InheritanceType All -AutoMapping:$true
    } else {
        Write-Verbose "This user does not have a manage defined in AD. You will have to manually delegate this mailbox."
    }

    # Set Strong Password
    Write-Verbose "Setting a 120 character strong password on the user account"

    $ascii=$NULL;
    For ($a=48;$a –le 122;$a++) {$ascii+=,[char][byte]$a }
    $PW= Get-TempPassword –length 120 –sourcedata $ascii
    $SecurePW = ConvertTo-SecureString $PW -asplaintext -force
    Set-ADAccountPassword -Identity $identity -NewPassword $SecurePW

    # Move User Account to the "Comapny - Disabled Accounts" OU
    Write-Verbose "Moving user account to the 'Comapny - Disabled Accounts' OU in AD"
    $OU = Get-ADOrganizationalUnit -filter * | where DistinguishedName -like "OU=Company- Disabled Accounts*" | select -ExpandProperty DistinguishedName
    Move-ADObject -Identity $DN -TargetPath $OU

    # Remove group memberships
    Write-Verbose "Removing all AD group memberships"
    $groups = Get-ADUser $identity -Properties MemberOf | select -ExpandProperty MemberOf
    foreach ($group in $groups) {
        Remove-ADGroupMember -Identity $group -Members $identity -Confirm:$false
    }

    # Disable AD Account 
    Write-Verbose "Disabling AD account"
    Disable-ADAccount $identity

    # Expire Account
    Write-Verbose "Setting AD account expiration"
    Set-ADAccountExpiration $identity -DateTime (get-date)


    Connect-MsolService -CurrentCredential
    # Remove Office 365 Licenses
    Write-Verbose "Removing the Users Office 365 Licenses"
    $Licenses = get-msoluser -UserPrincipalName $UserPrincipalName |
    select -ExpandProperty Licenses | 
    select -ExpandProperty AccountSkuId

    foreach ($License in $licenses) {
        Set-MsolUserLicense -UserPrincipalName $UserPrincipalName -RemoveLicenses $License
    }

    # Blocking Office 365 Logons
    Write-Verbose "Blocking the Users Office 365 Logons"
    Set-MsolUser -UserPrincipalName $UserPrincipalName -BlockCredential:$true

    # Forcing a sync between domain controllers
    Write-Verbose "Forcing a sync between domain controllers"
    $DC = Get-ADDomainController | select -ExpandProperty HostName
    Invoke-Command -ComputerName $DC -ScriptBlock {repadmin /syncall}
    Start-Sleep 30

    # Starting Sync From AD to Office 365 & Azure AD
    Write-Verbose 'Starting Sync From AD to Office 365 & Azure AD'
    Invoke-Command -ComputerName $DirSyncServer -ScriptBlock {Start-ScheduledTask 'Azure AD Sync Scheduler'}
    Start-Sleep 30
    Write-Verbose 'Complete!'
}
