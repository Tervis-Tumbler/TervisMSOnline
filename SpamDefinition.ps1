$SpamDomainDefinition = [PSCustomObject][Ordered]@{
    Domain = 
@"
karyatechsolutions.com
erpmaestro.com
softwareleadsusa.com
flycastpartners.com
fastlanetrainingus.com
xduce.com
hortonworks.com
5cok.com
liaison.com
cipher.com
turbonomic.com
aniruk.org
mitrend.com
email.bridgestonegolf.com
sampark.gov.in
arkadin.com
spencertech.com
rkwashburn.com
knowbe4.com
"@ -split "`r`n"
    EmailAddressesToTakeDomainFrom = 
@"
Quest International Users Group <Service@response.questdirect.org>
Brandon Lacanaria <blacanaria@lifesize.com>
Thomas Koll <tk@laplinkemail.com>
virtualizationwebinars <news@virtualizationwebinars.com>
Nexsan <connect@nexsan.com>
Digital Juice <enewsletter@digitaljuice.com>
"@ -split "`r`n"
    Reason = "Unsolicited email"
}