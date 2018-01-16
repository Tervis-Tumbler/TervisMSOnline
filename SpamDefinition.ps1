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
Monica Salazar <monica.salazar@zoom.us>
Michael Decker <mdecker@cesfl.com>
Megan from Metalogix <mwebb@reply.metalogix.com>
Training <Training@pmrgi.com>
Ziff Davis on behalf of Epson <elr_epson@updates.ziffdavisresearch.com>
Anooj Kumar <akumar@evoketechnologies.com>
Mike McMillan <mmcmillan@dsm.net>
Kayla Silverstein <KSilverstein@sonomapartners.com>
Talend <newsletter@talend.com>
CenturyLink Business <editor@centurylink-business.com>
Andrew Bale <marketing@tango-networks.com>
Kaseya <kaseyanews@kaseya.com>
Matt Bender <mbender@mainstreetdbas.com>
Ken Candela <kenc@kynetictech.com>
Information Management <msgs@product.information-management.com>
"@ -split "`r`n"
    Reason = "Unsolicited email"
}