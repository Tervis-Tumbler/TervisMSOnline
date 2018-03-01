#Requires -Version 5
$ModulePath = (Get-Module -ListAvailable TervisMSOnline).ModuleBase
. $ModulePath\SpamDefinition.ps1

function Test-TervisUserHasOffice365SharedMailbox {
    param(
        [parameter(mandatory)]$Identity
    )
    Import-TervisOffice365ExchangePSSession
    
    $UserPrincipalName = Get-ADUser -Identity $Identity | Select -ExpandProperty UserPrincipalName
    if (Get-O365Mailbox $UserPrincipalName -RecipientTypeDetails Shared -ErrorAction SilentlyContinue) {
        $true
    }
    else {
        $false
    }
}

function Import-TervisOffice365ExchangePSSession {
    Connect-TervisMsolService
    $MSOLUser = Get-MsolUser -UserPrincipalName "$env:USERNAME@$env:USERDOMAIN.com"
    
    if ($MSOLUser.StrongAuthenticationRequirements -and $MSOLUser.StrongAuthenticationRequirements.State -ne "Disabled") {
        Get-Module | Where-Object Name -Match tmp | Remove-Module -Force
        
        Get-PsSession |
        Where ComputerName -eq "outlook.office365.com" |
        Where ConfigurationName -eq "Microsoft.Exchange" |
        Remove-PSSession

        Import-TervisEXOPSSession
    } 
    else {
        Import-TervisMSOnlinePSSession
    }
}

function Import-TervisMSOnlinePSSession {
    [CmdletBinding()]
    param ()

    $Sessions = Get-PsSession |
    Where ComputerName -eq "ps.outlook.com" |
    Where ConfigurationName -eq "Microsoft.Exchange"
    
    $Sessions |
    Where State -eq "Broken" |
    Remove-PSSession

    $Session = $Sessions |
    Where State -eq "Opened" |
    Select -First 1

    if (-Not $Session) {
        $FunctionInfo = Get-Command Get-O365Mailbox -ErrorAction SilentlyContinue
        if ($FunctionInfo) {
            Remove-Module -Name $FunctionInfo.ModuleName            
        }
        Write-Verbose "Connect to Exchange Online"
        $Credential = Get-ExchangeOnlineCredential
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -Authentication Basic -ConnectionUri https://ps.outlook.com/powershell -AllowRedirection:$true -Credential $credential -WarningAction SilentlyContinue 
    }

    $FunctionInfo = Get-Command Get-O365Mailbox -ErrorAction SilentlyContinue
    if (-not $FunctionInfo) {
        Import-Module (Import-PSSession $Session -DisableNameChecking -AllowClobber) -DisableNameChecking -Global -Prefix "O365"
    }
}

function Import-TervisEXOPSSession {
    $Sessions = Get-PsSession |
    Where ComputerName -eq "outlook.office365.com" |
    Where ConfigurationName -eq "Microsoft.Exchange"
    
    $Sessions |
    Where State -eq "Broken" |
    Remove-PSSession

    $Session = $Sessions |
    Where State -eq "Opened" |
    Select -First 1

    if (-Not $Session) {
        $ExoScriptPath = Get-ExoPSSessionScriptPath
        Import-Module $ExoScriptPath -Force
        Connect-EXOPSSession -UserPrincipalName "$env:USERNAME@$env:USERDOMAIN.com" | Out-Null

        $Session = Get-PsSession |
        Where ComputerName -eq "outlook.office365.com" |
        Where ConfigurationName -eq "Microsoft.Exchange" |
        Where State -eq "Opened" |
        Select -First 1
    }
    
    $FunctionInfo = Get-Command Get-O365Mailbox -ErrorAction SilentlyContinue
    if (-not $FunctionInfo) {
        Import-Module (Import-PSSession $Session -DisableNameChecking -AllowClobber) -DisableNameChecking -Global -Prefix "O365"
    }
}

function Connect-TervisMsolService {
    try {
        Get-MsolDomain -ErrorAction Stop | Out-Null
    }
    catch {
        Connect-MsolService
    }
}

function Get-ExchangeOnlineCredential {
    Import-Clixml $env:USERPROFILE\ExchangeOnlineCredential.txt
}

function Set-TervisMSOLUserLicense {
    param (
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$UserPrincipalName,
        [ValidateSet("E3","E1")][Parameter(Mandatory)]$License
    )
    process {
        Connect-TervisMsolService
        
        $EnterprisePackSKU = Get-MsolAccountSku | 
        Where-Object {$_.AccountSkuID -match "ENTERPRISEPACK"} |
        Select-Object -ExpandProperty AccountSkuID

        $StandardPackSKU = Get-MsolAccountSku | 
        Where-Object {$_.AccountSkuID -match "STANDARDPACK"} |
        Select-Object -ExpandProperty AccountSkuID

        if ($License -eq "E3") { 
            $PackSkuToAdd = $EnterprisePackSKU
            $PackSkuToRemove = $StandardPackSKU
        } elseif ($License -eq "E1") { 
            $PackSkuToAdd = $StandardPackSKU
            $PackSkuToRemove = $EnterprisePackSKU
        }
        
        $AvailableLicenses = Get-MsolAccountSku | Where AccountSkuID -eq $PackSkuToAdd
            
        if ($AvailableLicenses.ConsumedUnits -ge $AvailableLicenses.ActiveUnits) {
            Throw "There are not any $License licenses available to assign to this user."
        }

        $MSOLUser = Get-MsolUser -UserPrincipalName $UserPrincipalName
        
        $CurrentLicenses = $MSOLUser.Licenses |
        Where-Object AccountSkuID -in $EnterprisePackSKU,$StandardPackSKU
        
        if ($CurrentLicenses.AccountSkuId -notcontains $PackSkuToAdd) {
            Set-MsolUser -UserPrincipalName $UserPrincipalName -UsageLocation 'US'
            Set-MsolUserLicense -UserPrincipalName $UserPrincipalName -AddLicenses $PackSkuToAdd
        }
        
        if ($CurrentLicenses.AccountSkuId -contains $PackSkuToRemove) {
            Set-MsolUserLicense -UserPrincipalName $UserPrincipalName -RemoveLicenses $PackSkuToRemove
        }
    }
}

function Set-ExchangeOnlineCredential {
    param(
        [Parameter(Mandatory)][System.Management.Automation.PSCredential]$Credential
    )
    $Credential | Export-Clixml $env:USERPROFILE\ExchangeOnlineCredential.txt
}

function Install-TervisMSOnline {
    param(
        [System.Management.Automation.PSCredential]$ExchangeOnlineCredential
    )
    <# 
    You must install the "Microsoft Online Services Sign-In Assistant for IT Professionals RTW" and 
    the "Azure Active Directory Module for Windows PowerShell (64-bit version)" before you can run this.
    The links are below.
    http://go.microsoft.com/fwlink/?LinkID=286152
    http://go.microsoft.com/fwlink/p/?linkid=236297
    #>
    if (-not $ExchangeOnlineCredential) {
        get-credential -message "Please supply the credentials to access ExchangeOnline. Username must be in the form UserName@Domain.com"
    }

    Install-Module -Name MSOnline
    # Below is depricated but we haven't figured out all the correct versions of things currently
    # Set-ExchangeOnlineCredential -Credential $ExchangeOnlineCredential 
    # Write-Verbose -Message "Installing Microsoft Online Services Sign-In Assistant for IT Professionals RTW..."
    # Install-TervisChocolateyPackageInstall -PackageName msonline-signin-assistant
    
    # Write-Verbose -Message "Installing Azure Active Directory Module for Windows PowerShell (64-bit version)..."
    # Install-TervisChocolateyPackageInstall -PackageName azure-ad-powershell-module
}

function Remove-TervisMobileDevice {
    [CmdletBinding()]
    param(
        [parameter(mandatory)]$Identity
    )
    Import-TervisOffice365ExchangePSSession
    Write-Verbose "Office 365 Removing Mobile Devices"    
    
    $UserPrincipalName = Get-ADUser -Identity $Identity -properties UserPrincipalName -ErrorAction SilentlyContinue |
    Select-Object -ExpandProperty UserPrincipalName

    Get-O365MobileDevice -Mailbox $UserPrincipalName -ErrorAction SilentlyContinue | 
    Write-VerboseAdvanced -PassThrough -Verbose:($VerbosePreference -ne "SilentlyContinue") |
    Remove-O365MobileDevice -Confirm:$false
}

function Remove-TervisMSOLUser {
    [CmdletBinding()]
    param(
        [parameter(mandatory)]$Identity,
        $IdentityOfUserToReceiveAccessToRemovedUsersMailbox
    )

    $UserObject = get-aduser $Identity -properties DistinguishedName,UserPrincipalName,ProtectedFromAccidentalDeletion
    $DN = $UserObject | select -ExpandProperty DistinguishedName
    $UserPrincipalName = $UserObject | select -ExpandProperty UserPrincipalName

    Import-TervisOffice365ExchangePSSession

    Remove-TervisMobileDevice -Identity $Identity

    Write-Verbose "Convert the users mailbox to a shared mailbox"
    Set-O365Mailbox $UserPrincipalName -Type shared

    Write-Verbose "Granting mailbox permissions to the $IdentityOfUserToReceiveAccessToRemovedUsersMailbox"
    if ($IdentityOfUserToReceiveAccessToRemovedUsersMailbox) {
        Add-O365MailboxPermission -Identity $UserPrincipalName -User $IdentityOfUserToReceiveAccessToRemovedUsersMailbox -AccessRights FullAccess -InheritanceType All -AutoMapping:$true | Out-Null
    }
    
    Connect-TervisMsolService
    Write-Verbose "Removing the Users Office 365 Licenses"
    $Licenses = Get-MsolUser -UserPrincipalName $UserPrincipalName |
        select -ExpandProperty Licenses | 
        select -ExpandProperty AccountSkuId

    foreach ($License in $licenses) {
        Set-MsolUserLicense -UserPrincipalName $UserPrincipalName -RemoveLicenses $License
    }

    Write-Verbose "Blocking the Users Office 365 Logons"
    Set-MsolUser -UserPrincipalName $UserPrincipalName -BlockCredential:$true
}

function Send-SupervisorOfTerminatedUserSharedEmailInstructions {
    param(
      $UserNameOfSupervisor,  
      $UserNameOfTerminatedUser  
    )
    Write-Verbose "Sending instructions to supervisor for Outlook for Mac"
    $ADObjectOfSupervisor = Get-ADUser -Identity $UserNameOfSupervisor
    $FirstNameOfSupervisor = $ADObjectOfSupervisor.GivenName
    $EmailAddressofSupervisor = $ADObjectOfSupervisor.UserPrincipalName

    $ADObjectOfTerminatedUser = Get-ADUser -Identity $UserNameOfTerminatedUser
    $FullNameofTerminatedUser = $ADObjectOfTerminatedUser.Name

    $Outlook2011Instructions = "\\fs1\DisasterRecovery\Source Controlled Items\TervisMSOnline\Add Shared Mailbox to Outlook 2011 Mac.docx"
    $Outlook2016Instructions = "\\fs1\DisasterRecovery\Source Controlled Items\TervisMSOnline\Add Shared Mailbox to Outlook 2016 Mac.docx"

    $HTMLBody = @"
<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40"><head>
<meta http-equiv="Content-Type" content="text/html; charset=us-ascii">
<meta name="Generator" content="Microsoft Word 15 (filtered medium)">
<!--[if !mso]><style>v\:* {behavior:url(#default#VML);}
o\:* {behavior:url(#default#VML);}
w\:* {behavior:url(#default#VML);}
.shape {behavior:url(#default#VML);}
</style><![endif]--><style><!--
/* Font Definitions */
@font-face
	{font-family:"Cambria Math";
	panose-1:2 4 5 3 5 4 6 3 2 4;}
@font-face
	{font-family:Calibri;
	panose-1:2 15 5 2 2 2 4 3 2 4;}
@font-face
	{font-family:Verdana;
	panose-1:2 11 6 4 3 5 4 4 2 4;}
/* Style Definitions */
p.MsoNormal, li.MsoNormal, div.MsoNormal
	{margin:0in;
	margin-bottom:.0001pt;
	font-size:11.0pt;
	font-family:"Calibri",sans-serif;}
a:link, span.MsoHyperlink
	{mso-style-priority:99;
	color:#0563C1;
	text-decoration:underline;}
a:visited, span.MsoHyperlinkFollowed
	{mso-style-priority:99;
	color:#954F72;
	text-decoration:underline;}
span.EmailStyle17
	{mso-style-type:personal-compose;
	font-family:"Calibri",sans-serif;
	color:windowtext;}
.MsoChpDefault
	{mso-style-type:export-only;
	font-family:"Calibri",sans-serif;}
@page WordSection1
	{size:8.5in 11.0in;
	margin:1.0in 1.0in 1.0in 1.0in;}
div.WordSection1
	{page:WordSection1;}
--></style><!--[if gte mso 9]><xml>
<o:shapedefaults v:ext="edit" spidmax="1026" />
</xml><![endif]--><!--[if gte mso 9]><xml>
<o:shapelayout v:ext="edit">
<o:idmap v:ext="edit" data="1" />
</o:shapelayout></xml><![endif]-->
</head>
<body lang="EN-US" link="#0563C1" vlink="#954F72">
<div class="WordSection1">
<p class="MsoNormal">$FirstNameOfSupervisor,<o:p></o:p></p>
<p class="MsoNormal"><o:p>&nbsp;</o:p></p>
<p class="MsoNormal">You have been given Full Access permission to the mailbox of:&nbsp; $FullNameofTerminatedUser<o:p></o:p></p>
<p class="MsoNormal"><o:p>&nbsp;</o:p></p>
<p class="MsoNormal">Please find the attached instructions on how to attach this shared mailbox in Outlook.&nbsp; You will see that there are two documents with instructions based on which version of Outlook that you have.<o:p></o:p></p>
<p class="MsoNormal"><o:p>&nbsp;</o:p></p>
<p class="MsoNormal">If you experience any issues, please contact the Help Desk at 2248 or externally at (941) 441-3168.<o:p></o:p></p>
<p class="MsoNormal"><o:p>&nbsp;</o:p></p>
<p class="MsoNormal">Thank you,<o:p></o:p></p>
<p class="MsoNormal"><o:p>&nbsp;</o:p></p>
<table class="MsoNormalTable" border="0" cellspacing="0" cellpadding="0" width="485" style="width:363.65pt;border-collapse:collapse">
<tbody>
<tr style="height:29.95pt">
<td width="447" valign="top" style="width:335.45pt;padding:0in 0in 0in 0in;height:29.95pt">
<p class="MsoNormal" style="line-height:115%"><span style="font-size:10.0pt;line-height:115%;font-family:&quot;Verdana&quot;,sans-serif;color:#595959">HELP DESK TEAM</span><span style="font-size:10.0pt;line-height:115%;font-family:&quot;Verdana&quot;,sans-serif;color:#595959"><o:p></o:p></span></p>
<p class="MsoNormal" style="line-height:115%"><span style="font-size:10.0pt;line-height:115%;font-family:&quot;Verdana&quot;,sans-serif;color:#595959">d: 2248 or 941-441-3168<o:p></o:p></span></p>
<p class="MsoNormal" style="line-height:115%"><span style="font-size:10.0pt;line-height:115%"><img width="176" height="61" style="width:1.8333in;height:.6354in" id="Picture_x0020_25" src="https://sharepoint.tervis.com/SiteCollectionImages/NEW_Logo.jpg" alt="Tervis_Color_Logo_URL"><o:p></o:p></span></p>
<p class="MsoNormal" style="margin-left:4.5pt;line-height:115%"><span style="font-size:10.0pt;line-height:115%"><o:p>&nbsp;</o:p></span></p>
</td>
<td width="38" valign="top" style="width:28.2pt;padding:0in 5.4pt 0in 5.4pt;height:29.95pt">
<p class="MsoNormal" align="center" style="margin-left:-23.4pt;text-align:center;line-height:115%">
<o:p>&nbsp;</o:p></p>
</td>
</tr>
</tbody>
</table>
<p class="MsoNormal"><o:p>&nbsp;</o:p></p>
<p class="MsoNormal"><o:p>&nbsp;</o:p></p>
</div>
CONFIDENTIALITY NOTICE: At Tervis we make great drinkware that helps people celebrate great moments. Sometimes we also make mistakes and send emails to the wrong address. If you received this in error, please don&#8217;t read or pass it on, as it may contain confidential
 and/or privileged information and is intended only for the recipient(s) to which it is addressed. Any other use is strictly prohibited. Please notify the sender so that we may correct our internal records and then delete the original message. Thanks.
</body>
</html>
"@

    Send-TervisMailMessage -To $EmailAddressofSupervisor -From 'Help Desk Team <HelpDeskTeam@tervis.com>' -Subject 'Instructions to Add Shared Email to Outlook' -Body $HTMLBody -Attachments $Outlook2011Instructions, $Outlook2016Instructions -bodyashtml
}

function Move-SharedMailboxObjects {
    param(
        [parameter(mandatory)]$DistinguishedNameOfTargetOU
    )
    Write-Verbose "Connect to Exchange Online"
    $Sessions = Get-PsSession
    $Connected = $false
    Foreach ($Session in $Sessions) {
        if ($Session.ComputerName -eq 'ps.outlook.com' -and $Session.ConfigurationName -eq 'Microsoft.Exchange' -and $Session.State -eq 'Opened') {
            $Connected = $true
        } elseif ($Session.ComputerName -eq 'ps.outlook.com' -and $Session.ConfigurationName -eq 'Microsoft.Exchange' -and $Session.State -eq 'Broken') {
            Remove-PSSession $Session
        }
    }
    if ($Connected -eq $false) {
        Write-Verbose "Connect to Exchange Online"
        $Credential = Get-ExchangeOnlineCredential
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -Authentication Basic -ConnectionUri https://ps.outlook.com/powershell -AllowRedirection:$true -Credential $credential
        Import-PSSession $Session -Prefix 'O365' -DisableNameChecking -AllowClobber
    }

    $SharedMailboxOnOffice365 = Get-O365Mailbox -RecipientTypeDetails 'Shared' -ResultSize 'Unlimited'
    foreach ($SharedMailbox in $SharedMailboxOnOffice365) {
        [string]$UserName = $SharedMailbox | Select -ExpandProperty UserPrincipalName
        $ADUser = Get-ADObject -Filter {UserPrincipalName -eq $UserName} 
        if (-NOT ($ADUser.DistinguishedName -match $DistinguishedNameOfTargetOU)) {
            $ADUser | Set-ADObject -ProtectedFromAccidentalDeletion $false
            $ADUser | Move-ADObject -TargetPath $DistinguishedNameOfTargetOU -Confirm:$false
            $ADUser = Get-ADObject -Filter {UserPrincipalName -eq $UserName} 
            $ADUser | Set-ADObject -ProtectedFromAccidentalDeletion $true
        }
    }
}

function Install-MoveSharedMailboxObjectsScheduledTasks {
    param (
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$ComputerName
    )
    begin {
        $ScheduledTaskCredential = New-Object System.Management.Automation.PSCredential (Get-PasswordstateCredential -PasswordID 259)
        $TargetOU = Get-ADOrganizationalUnit -Filter {Name -eq "Shared Mailbox"} | `
            Where DistinguishedName -match 'OU=Shared Mailbox,OU=Exchange,DC=' | `
            Select -ExpandProperty DistinguishedName
        $Execute = 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe'
        $Argument = "-Command `"& {Move-SharedMailboxObjects -DistinguishedNameOfTargetOU `'$TargetOU`'}`""
    }
    process {
        $CimSession = New-CimSession -ComputerName $ComputerName
        If (-NOT (Get-ScheduledTask -TaskName Move-SharedMailboxObjects -CimSession $CimSession -ErrorAction SilentlyContinue)) {
            Install-TervisScheduledTask -Credential $ScheduledTaskCredential -TaskName Move-SharedMailboxObjects -Execute $Execute -Argument $Argument -RepetitionIntervalName EveryDayAt2am -ComputerName $ComputerName
        }
    }
}

function Add-TervisMSOnlineAdminRoleMember {
    [CmdletBinding()]
    param (
        [parameter(mandatory)]$UserPrincipalName,
        [ValidateSet(
        "Billing Administrator", 
        "Company Administrator", 
        "Compliance Administrator", 
        "Device Administrators", 
        "Exchange Service Administrator", 
        "Helpdesk Administrator",
        "Lync Service Administrator",
        "Privileged Role Administrator",
        "Security Administrator",
        "Service Support Administrator",
        "SharePoint Service Administrator",
        "User Account Administrator"
        )]$RoleName
    )
    
    Connect-TervisMsolService

    Add-MsolRoleMember -RoleMemberEmailAddress $UserPrincipalName -RoleName $RoleName
}

function Set-DTCNewHireO365MailboxPermissions{
    param(
    	[cmdletbinding()]
        [parameter(mandatory)]$User
    )
    Import-TervisOffice365ExchangePSSession
    
    $Mailboxes = "customercare","weborderstatus","webreturns","customyzer"
    
    foreach ($Mailbox in $Mailboxes){
        Add-O365MailboxPermission -Identity $Mailbox -User $User -AccessRights "FullAccess"
        Add-O365RecipientPermission -Identity $Mailbox -AccessRights SendAs -Trustee $User -Confirm:$false
    }
}

function Add-TervisO365MailboxPermission{
    param(
        [cmdletbinding()]
        [parameter(Mandatory)]$User,
        [parameter(Mandatory)]$Mailbox,
        [ValidateSet("FullAccess","SendAs","FullAccessAndSendAs")][parameter(Mandatory)]$Permission
    )
    Import-TervisOffice365ExchangePSSession

    if ($Permission -eq "FullAccess"){
        Add-O365MailboxPermission -Identity $Mailbox -User $User -AccessRights "FullAccess"
    } elseif ($Permission -eq "SendAs"){
        Add-O365RecipientPermission -Identity $Mailbox -AccessRights SendAs -Trustee $User -Confirm:$false
    } elseif ($Permission -eq "FullAccessAndSendAs") {
        Add-O365MailboxPermission -Identity $Mailbox -User $User -AccessRights "FullAccess"
        Add-O365RecipientPermission -Identity $Mailbox -AccessRights SendAs -Trustee $User -Confirm:$false
    }
}

function Remove-TervisO365MailboxPermission{
    param(
        [cmdletbinding()]
        [parameter(Mandatory)]$User,
        [parameter(Mandatory)]$Mailbox,
        [ValidateSet("FullAccess","SendAs","FullAccessAndSendAs")][parameter(Mandatory)]$Permission
    )
    Import-TervisOffice365ExchangePSSession

    if ($Permission -eq "FullAccess"){
        Remove-O365MailboxPermission -Identity $Mailbox -User $User -AccessRights "FullAccess"
    } elseif ($Permission -eq "SendAs"){
        Remove-O365RecipientPermission -Identity $Mailbox -AccessRights SendAs -Trustee $User -Confirm:$false
    } elseif ($Permission -eq "FullAccessAndSendAs") {
        Remove-O365MailboxPermission -Identity $Mailbox -User $User -AccessRights "FullAccess"
        Remove-O365RecipientPermission -Identity $Mailbox -AccessRights SendAs -Trustee $User -Confirm:$false
    }
}

# Do Not Use. New-TervisMSOLUser is not ready for use with current hybrid setup with on-promise Exchange2016.
#function New-TervisMSOLUser{
#    [CmdletBinding()]
#    param(
#        [paramete r(mandatory)]$Identity,
#        [parameter(mandatory)]$AzureADConnectComputerName
#    )
#
#    $UserObject = get-aduser $Identity -properties DistinguishedName,UserPrincipalName,ProtectedFromAccidentalDeletion
#    $DN = $UserObject | select -ExpandProperty DistinguishedName
#    $UserPrincipalName = $UserObject | select -ExpandProperty UserPrincipalName
#
#    Import-TervisMSOnlinePSSession
#
#    Enable-O365Mailbox -Identity $Identity -RoleAssignmentPolicy "Default Role Assignment Policy"
#
#    $Credential = Get-ExchangeOnlineCredential
#    Connect-MsolService -Credential $Credential
#    Write-Verbose "Adding the Users Office 365 Licenses"
#    Set-MsolUserLicense -UserPrincipalName $UserPrincipalName -AddLicenses "Office 365 Enterprise E3"
#}

function Enable-Office365MultiFactorAuthentication {
    param (
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$UserPrincipalName
    )
    begin {
        Connect-TervisMsolService
    }
    process {
        $auth = New-Object -TypeName Microsoft.Online.Administration.StrongAuthenticationRequirement
        $auth.RelyingParty = "*"
        $auth.State = "Enforced"
        $auth.RememberDevicesNotIssuedBefore = (Get-Date)

        Set-MsolUser -UserPrincipalName $UserPrincipalName -StrongAuthenticationRequirements $auth
    }
}

function Disable-Office365MultiFactorAuthentication {
    param (
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$UserPrincipalName
    )
    begin {
        Connect-TervisMsolService
    }
    process {   
        $auth = New-Object -TypeName Microsoft.Online.Administration.StrongAuthenticationRequirement
        $auth.RelyingParty = "*"
        $auth.State = "Disabled"
        $auth.RememberDevicesNotIssuedBefore = (Get-Date)

        Set-MsolUser -UserPrincipalName $UserPrincipalName -StrongAuthenticationRequirements $auth
    }
}

function Get-MsolUsersWithAnE1OrE3LicenseExcludingServiceAccounts {
    param (
        [Switch]$ExcludeUsersWithStrongAuthenticationEnforced
    )
    Connect-TervisMsolService
    $AllMSOL = Get-MsolUser -All
    if ($ExcludeUsersWithStrongAuthenticationEnforced) {
        $MSOLUsersToFilter = $AllMSOL | where {$_.StrongAuthenticationRequirements.State -NE "Enforced"}
    } else {
        $MSOLUsersToFilter = $AllMSOL
    }

    $MSOLUsersToFilter | 
        where {
            $_.licenses.AccountSkuID -match "tervis0:ENTERPRISEPACK" -or
            $_.licenses.AccountSkuID -match "tervis0:STANDARDPACK"
        } | 
        where UserPrincipalName -NotMatch TTC_ |
        sort DisplayName | 
        select DisplayName, UserPrincipalName, Licenses
}

function Get-MsolUsersWithStrongAuthenticationNotConfigured {
    Connect-TervisMsolService
    $AllMsolUsers = Get-MsolUser -All | 
        where {
            $_.licenses.AccountSkuID -match "tervis0:ENTERPRISEPACK" -or
            $_.licenses.AccountSkuID -match "tervis0:STANDARDPACK"
        } | 
        where UserPrincipalName -NotMatch TTC_
    $MsolUsersWithStrongAuthenticationConfigured = $AllMsolUsers | where {$_.StrongAuthenticationMethods -ne $null}
    $MsolUsersWithStrongAuthenticationNotConfigured = $AllMsolUsers | where {$_.StrongAuthenticationMethods -eq $null}
    $MsolUsersWithStrongAuthenticationNotConfigured

    $AllUserCount = $AllMsolUsers.Count
    $ConfiguredUserCount = $MsolUsersWithStrongAuthenticationConfigured.Count
    $PercentageConfigured = "{0:N2}" -f ($ConfiguredUserCount * 100/$AllUserCount)
    Write-Warning "`nUsers with MFA configured:`t`t$ConfiguredUserCount`nTotal number of users:`t`t`t$AllUserCount`nPercent with MFA configured:`t$PercentageConfigured"
}

function Get-ExoPSSessionScriptPath {
    $RootPath = "$env:LOCALAPPDATA\Apps\2.0"
    $ExoScriptName = "CreateExoPSSession.ps1"
    $DllItemName = "Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
    $ExoScriptItems = Get-ChildItem -Path $RootPath -Name $ExoScriptName -Recurse
    $DllItems = Get-ChildItem -Path $RootPath -Name $DllItemName -Recurse

    foreach ($ExoScriptItem in $ExoScriptItems) {
        foreach ($DllItem in $DllItems) {
            if (($ExoScriptItem | Split-Path -Parent) -eq ($DllItem | Split-Path -Parent)) {
                $RealScriptItem = $ExoScriptItem
                break;
            }
        }
    }
    "$RootPath\$RealScriptItem"
}

function Invoke-ExoPSSessionScript {
    $ExoScriptPath = Get-ExoPSSessionScriptPath
    . $ExoScriptPath
}

function Sync-SpamDomainDefinitionWithOffice365 {
    Import-TervisEXOPSSession
    Set-O365HostedContentFilterPolicy -Identity default -BlockedSenderDomains @{Add=$SpamDomainDefinition.Domain}
    $DomainsFromEmailAddresses = $SpamDomainDefinition.EmailAddressesToTakeDomainFrom |
    ForEach-Object {         
        (
            $_ -split "@" |
            Select-Object -Skip 1
        ) -replace ">",""
    }
    if ($DomainsFromEmailAddresses) {
        Set-O365HostedContentFilterPolicy -Identity default -BlockedSenderDomains @{Add=$DomainsFromEmailAddresses}
    }
}

function Get-MsolUsersByLicenseType {
    param (
        [Parameter(Mandatory)]
        [ValidateSet("K1","E1","E3","Unlicensed")]
        $LicenseType
    )
    
    switch ($LicenseType) {
        "K1" {$License = "tervis0:EXCHANGEDESKLESS"}
        "E1" {$License = "tervis0:STANDARDPACK"}
        "E3" {$License = "tervis0:ENTERPRISEPACK"}
    }
    
    Connect-TervisMsolService
    $AllMSOL = Get-MsolUser -All
    
    if ($LicenseType -eq "Unlicensed") {
        $AllMSOL | where IsLicensed -eq $false
    } else {
        $AllMSOL | where {$_.Licenses.AccountSkuId -contains $License}
    }
}

function Get-TervisO365MailboxUserHasAccessto {
    param ([parameter(Mandatory)]$User,
          [Parameter(Mandatory)][ValidateSet("SharedMailBox","UserMailbox")][String]$mailboxType 
    )
    Import-TervisOffice365ExchangePSSession
    Get-O365Mailbox -RecipientTypeDetails $mailboxType -ResultSize unlimited | Get-O365MailboxPermission -User $User 

}

function Connect-EXOPSSessionWithinExchangeOnlineShell {
    if (Get-ChildItem -Path function:\Connect-EXOPSSession -ErrorAction SilentlyContinue) { #Test whether we are in a Exchange Online PowerShell Module console shell
        $GetMailboxFunction = Get-ChildItem -Path function:\Get-Mailbox #Test whether we have already ran Connect-EXOPSSession
        if ( -not $GetMailboxFunction ) {
            $EXOPsSessionModule = Connect-EXOPSSession #Store the tmp module output in case we want to try later to import with namespace
        }
        $true #We were in an Exchange Online PowerShell Module console and we have imported the functions 
    } else {
        $false #We were not in an Exchange Online PowerShell Module console
    }
}