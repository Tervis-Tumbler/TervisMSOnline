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

function Test-TervisUserHasMSOnlineMailbox {
    param(
        [parameter(mandatory)]$Identity
    )
    write-verbose "Connect to Exchange Online with your user@domain.com credentials"
    $WarningPreference = 'SilentlyContinue'
    $Credential = Import-Clixml $env:USERPROFILE\ExchangeOnlineCredential.txt
    $CloudSession = New-PSSession -Name MSOnlineSession -ConfigurationName Microsoft.Exchange -Authentication Basic -ConnectionUri https://ps.outlook.com/powershell -AllowRedirection:$true -Credential $credential
    
    Import-PSSession $CloudSession -AllowClobber | Out-Null
    
    if (Get-Mailbox $Identity -ErrorAction SilentlyContinue) {
        $MsolMailbox = $true
    }
    else {
        $MSolMailbox = $false    
    }
    
    $MsolMailbox
    
    Remove-PSSession -Name MSOnlineSession
}

function Test-TervisUserHasOnPremMailbox {
    param(
        [parameter(mandatory)]$Identity
    )
    $WarningPreference = 'SilentlyContinue'
    $OnPremSession = New-PSSession -Name OnPremSession -ConfigurationName Microsoft.Exchange -Authentication Kerberos -ConnectionUri http://exchange2010.tervis.prv/powershell
    
    Import-PSSession $OnPremSession -AllowClobber | Out-Null
    
    if (Get-Mailbox $Identity -ErrorAction SilentlyContinue) {
        $OnPremMailbox = $true
    }
    else {
        $OnPremMailbox = $false    
    }
    
    $OnPremMailbox
    
    Remove-PSSession -Name OnPremSession
}

function Import-TervisMSOnlinePSSession {
    [CmdletBinding()]
    param ()
    write-verbose "Connect to Exchange Online with your user@domain.com credentials"
    $Credential = Import-Clixml $env:USERPROFILE\ExchangeOnlineCredential.txt
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -Authentication Basic -ConnectionUri https://ps.outlook.com/powershell -AllowRedirection:$true -Credential $credential
    Import-PSSession $Session -DisableNameChecking
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
    Write-Verbose -Message "Installing Microsoft Online Services Sign-In Assistant for IT Professionals RTW..."
    Install-TervisChocolateyPackageInstall -PackageName msonline-signin-assistant -Force
    
    Write-Verbose -Message "Installing Azure Active Directory Module for Windows PowerShell (64-bit version)..."
    Install-TervisChocolateyPackageInstall -PackageName azure-ad-powershell-module -Force
}

function Remove-TervisMSOLUser{
    [CmdletBinding()]
    param(
        [parameter(mandatory)]$Identity,
        [parameter(mandatory)]$AzureADConnectComputerName,
        $IdentityOfUserToReceiveAccessToRemovedUsersMailbox
    )

    $UserObject = get-aduser $Identity -properties DistinguishedName,UserPrincipalName
    $DN = $UserObject | select -ExpandProperty DistinguishedName
    $UserPrincipalName = $UserObject | select -ExpandProperty UserPrincipalName

    Import-TervisMSOnlinePSSession

    Write-Verbose "Removing Users Active Sync Devices"
    if (Get-ActiveSyncDevice -Mailbox $UserPrincipalName -ErrorAction SilentlyContinue) {
        Get-ActiveSyncDevice -Mailbox $UserPrincipalName | Remove-ActiveSyncDevice
    }

    Write-Verbose "Convert the users mailbox to a shared mailbox"
    Set-Mailbox $UserPrincipalName -Type shared

    Write-Verbose "Granting mailbox permissions to the $IdentityOfUserToReceiveAccessToRemovedUsersMailbox"
    if ($IdentityOfUserToReceiveAccessToRemovedUsersMailbox) {
        Add-MailboxPermission -Identity $UserPrincipalName -User $IdentityOfUserToReceiveAccessToRemovedUsersMailbox -AccessRights FullAccess -InheritanceType All -AutoMapping:$true
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
    write-verbose "Connect to Exchange Online with your user@domain.com credentials"
    $Credential = Import-Clixml $env:USERPROFILE\ExchangeOnlineCredential.txt
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -Authentication Basic -ConnectionUri https://ps.outlook.com/powershell -AllowRedirection:$true -Credential $credential
    Import-PSSession $Session -Prefix 'Cloud' -DisableNameChecking

    $SharedMailboxOnOffice365 = Get-CloudMailbox -RecipientTypeDetails 'Shared' -ResultSize 'Unlimited'
    foreach ($SharedMailbox in $SharedMailboxOnOffice365) {
        [string]$UserName = ($SharedMailbox | Select -ExpandProperty UserPrincipalName).Split('@')[0]
        Get-ADUser $UserName | Move-ADObject -TargetPath $DistinguishedNameOfTargetOU -Confirm:$false
    }
}