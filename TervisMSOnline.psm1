#Requires -Modules MSOnline
#Requires -Version 5

Function Get-TempPassword() {
    Param(
        [int]$length=120,
        [string[]]$sourcedata
    )

    For ($loop=1; $loop -le $length; $loop++) {
        $TempPassword+=($sourcedata | Get-Random)
    }
    Return $TempPassword
}

Function Test-TervisUserHasMailbox {
    param(
        [parameter(mandatory)]$Identity
    )
    write-verbose "Connect to Exchange Online with your user@domain.com credentials"
    $Credential = Import-Clixml $env:USERPROFILE\ExchangeOnlineCredential.txt
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -Authentication Basic -ConnectionUri https://ps.outlook.com/powershell -AllowRedirection:$true -Credential $credential
    Import-PSSession $Session -Prefix Cloud -DisableNameChecking
    $MsolMailbox = $false
    $OnPremiseMailbox = $false
    add-pssnapin Microsoft.Exchange.Management.PowerShell.E2010
    if (Get-CloudMailbox $Identity -ErrorAction SilentlyContinue) {
        $MsolMailbox = $true
    } elseif (get-mailbox $Identity -ErrorAction SilentlyContinue){
        $OnPremiseMailbox = $true
    }
}

function Install-TervisMSOnline {
    param(
        [System.Management.Automation.PSCredential]$ExchangeOnlineCredential = $(get-credential -message "Please supply the credentials to access ExchangeOnline. Username must be in the form UserName@Domain.com")
    )
    <# 
    You must install the "Microsoft Online Services Sign-In Assistant for IT Professionals RTW" and 
    the "Azure Active Directory Module for Windows PowerShell (64-bit version)" before you can run this.
    The links are below.
    http://go.microsoft.com/fwlink/?LinkID=286152
    http://go.microsoft.com/fwlink/p/?linkid=236297
    #>

    $ExchangeOnlineCredential | Export-Clixml $env:USERPROFILE\ExchangeOnlineCredential.txt
    start-process "http://go.microsoft.com/fwlink/?LinkID=286152"
    Read-Host "Please install Microsoft Online Services Sign-In Assistant for IT Professionals RTW and then hit enter"
    Invoke-WebRequest -Uri http://go.microsoft.com/fwlink/p/?linkid=236297 -OutFile AdministrationConfig-en.msi
    Start-Process AdministrationConfig-en.msi
}

function Remove-TervisMSOLUser{
    param(
        [parameter(mandatory)]$Identity,
        [parameter(mandatory)]$AzureADConnectComputerName,
        $IdentityOfUserToRecieveAccessToRemovedUsersMailbox
    )

    $UserObject = get-aduser $Identity -properties DistinguishedName,UserPrincipalName
    $DN = $UserObject | select -ExpandProperty DistinguishedName
    $UserPrincipalName = $UserObject | select -ExpandProperty UserPrincipalName

    write-verbose "Connect to Exchange Online with your user@domain.com credentials"
    $Credential = Import-Clixml $env:USERPROFILE\ExchangeOnlineCredential.txt
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -Authentication Basic -ConnectionUri https://ps.outlook.com/powershell -AllowRedirection:$true -Credential $credential
    Import-PSSession $Session -DisableNameChecking

    Write-Verbose "Removing Users Active Sync Devices"
    if (Get-ActiveSyncDevice -Mailbox $UserPrincipalName -ErrorAction SilentlyContinue) {
        Get-ActiveSyncDevice -Mailbox $UserPrincipalName | Remove-ActiveSyncDevice
    }

    Write-Verbose "Convert the users mailbox to a shared mailbox"
    Set-Mailbox $UserPrincipalName -Type shared

    Write-Verbose "Granting mailbox permissions to the $IdentityOfUserToRecieveAccessToRemovedUsersMailbox"
    if ($IdentityOfUserToRecieveAccessToRemovedUsersMailbox) {
        Add-MailboxPermission -Identity $UserPrincipalName -User $IdentityOfUserToRecieveAccessToRemovedUsersMailbox -AccessRights FullAccess -InheritanceType All -AutoMapping:$true
    }

    Write-Verbose "Setting a 120 character strong password on the user account"

    $ascii=$NULL;
    For ($a=48;$a -le 122;$a++) {$ascii+=,[char][byte]$a }
    $PW= Get-TempPassword -length 120 -sourcedata $ascii
    $SecurePW = ConvertTo-SecureString $PW -asplaintext -force
    Set-ADAccountPassword -Identity $identity -NewPassword $SecurePW

    Write-Verbose "Moving user account to the 'Comapny - Disabled Accounts' OU in AD"
    $OU = Get-ADOrganizationalUnit -filter * | where DistinguishedName -like "OU=Company- Disabled Accounts*" | select -ExpandProperty DistinguishedName
    Move-ADObject -Identity $DN -TargetPath $OU

    Write-Verbose "Removing all AD group memberships"
    $groups = Get-ADUser $identity -Properties MemberOf | select -ExpandProperty MemberOf
    foreach ($group in $groups) {
        Remove-ADGroupMember -Identity $group -Members $identity -Confirm:$false
    }

    Write-Verbose "Disabling AD account"
    Disable-ADAccount $identity

    Write-Verbose "Setting AD account expiration"
    Set-ADAccountExpiration $identity -DateTime (get-date)


    Connect-MsolService -Credential $Credential
    Write-Verbose "Removing the Users Office 365 Licenses"
    $Licenses = get-msoluser -UserPrincipalName $UserPrincipalName |
    select -ExpandProperty Licenses | 
    select -ExpandProperty AccountSkuId

    foreach ($License in $licenses) {
        Set-MsolUserLicense -UserPrincipalName $UserPrincipalName -RemoveLicenses $License
    }

    Write-Verbose "Blocking the Users Office 365 Logons"
    Set-MsolUser -UserPrincipalName $UserPrincipalName -BlockCredential:$true

    Write-Verbose "Forcing a sync between domain controllers"
    $DC = Get-ADDomainController | select -ExpandProperty HostName
    Invoke-Command -ComputerName $DC -ScriptBlock {repadmin /syncall}
    Start-Sleep 30

    Write-Verbose 'Starting Sync From AD to Office 365 & Azure AD'
    Invoke-Command -ComputerName $AzureADConnectComputerName -ScriptBlock {Start-ScheduledTask 'Azure AD Sync Scheduler'}
    Start-Sleep 30
    Write-Verbose 'Complete!'
}
