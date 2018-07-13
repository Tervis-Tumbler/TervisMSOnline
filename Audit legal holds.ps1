Import-TervisEXOPSSession
Get-O365Mailbox -Filter {LitigationHoldEnabled -eq $True} | measure
Get-O365Mailbox -Filter {LitigationHoldEnabled -eq $false} | measure
$MailboxesWithoutLitigationHold = Get-O365Mailbox -Filter {LitigationHoldEnabled -eq $false}
$MailboxesWithoutLitigationHold.samaccountname
$MailboxesWithoutLitigationHold[0]
$MailboxesWithoutLitigationHold[0] | fl *
$MailboxesWithoutLitigationHold[0] | fl *shar*
$MailboxesWithoutLitigationHold | where {$_.IsShared} | Measure
$MailboxesWithoutLitigationHoldNotShared = $MailboxesWithoutLitigationHold | where {-not $_.IsShared} | Measure
$MailboxesWithoutLitigationHoldNotShared = $MailboxesWithoutLitigationHold | where {-not $_.IsShared}
$MailboxesWithoutLitigationHoldNotShared[0]
$MailboxesWithoutLitigationHoldNotShared[0] | FL *
$MailboxesWithoutLitigationHoldNotShared
Get-O365Mailbox -Filter * | select -ExpandProperty litigationholddate | Group
(Get-date)
(Get-date).Date
Get-O365Mailbox -Filter * | select -ExpandProperty litigationholddate | %{ $_.date} | Group
$Dates = Get-O365Mailbox -Filter * | select -ExpandProperty litigationholddate


$MailboxesWithoutLitigationHold = Get-O365Mailbox -Filter {LitigationHoldEnabled -eq $false}
$MailboxesWithoutLitigationHoldNotShared = $MailboxesWithoutLitigationHold | where {-not $_.IsShared}