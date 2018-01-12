$SpamDomainDefinition = [PSCustomObject][Ordered]@{
    Domain = 
@"
karyatechsolutions.com
erpmaestro.com
softwareleadsusa.com
flycastpartners.com
fastlanetrainingus.com
xduce.com
agates@hortonworks.com
"@ -split "`r`n"
    Reason = "Unsolicited email"
}