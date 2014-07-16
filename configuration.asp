<%
' ==============Global Configuration=============
' ===============================================
sShipperPostalCode = "91342"
sShipperCountryCode = "US"

sShippingMarkupFactor = 1
sShippingMarkupFlatRate = 0

sDisplayErrors = True
' ============End Global Configuration===========
' ===============================================



' =====================UPS=======================
' ===============================================
sUPSAccessLiscenseNumber = "xxxxxxxxxxxxxxx"
sUPSUserID = "xxxxxx"
sUPSPassword = "xxxxxx"

sUPSPickupType = "19"
' Valid values:
'  01 - part of a daily pickup
'  06 - a one time pickup 
'  19 - dropped off at a UPS Letter Center

' ===============================================
' ===============================================


' ====================FEDEX======================
' ===============================================
'These values are *required* to do transactions with FedEx.
sFEDEXAccountNumber = "xxxxxxxxxx"
sFEDEXMeterNumber = "xxxxxx"


' -----------------Meter Number------------------
'These values must be passed when you first request your meter number.
'After you have a valid meter number, they are never used again.
sFEDEXName = "Your Name"
sFEDEXCompanyName = "xxxxxxxxx"
sFEDEXPhoneNumber = "xxxxxxxxxxx"
sFEDEXEmailAddress = "xxxxxxx@xxxxxx.com"
sFEDEXAddress = "xxxxxxxxxxxxxxxxxxxxxx"
sFEDEXCity = "xxxxxxxxx"
sFEDEXState = "xx"
' -----------------------------------------------

' -=FedEx Services Configuration=-
'Here you can configure the FedEx services you'd like to enable. Some services don't or won't apply
'to your situation so disabling the one's you're not interested in will improve performance since
'each service queried requires a seperate call to FedEx to retrieve the rate.
sEnabledFedExServices = "PRIORITYOVERNIGHT, STANDARDOVERNIGHT, FIRSTOVERNIGHT, FEDEX2DAY, FEDEXEXPRESSSAVER, FEDEXGROUND, GROUNDHOMEDELIVERY, INTERNATIONALPRIORITY, INTERNATIONALECONOMY, INTERNATIONALFIRST, FEDEX1DAYFREIGHT, FEDEX2DAYFREIGHT, FEDEX3DAYFREIGHT, INTERNATIONALPRIORITY FREIGHT, INTERNATIONALECONOMY FREIGHT, EUROPEFIRSTINTERNATIONALPRIORITY"

'Moving the services to this variable simply allow you to keep track of the one's that are disabled.
'this string is currently not used for anything but may be utilized later.
sDisabledFedExServices = ""

' -= Drop Off Type =-
sFedExDropOffType = "REGULARPICKUP"

' Valid values:
' REQUESTCOURIER
' DROPBOX
' BUSINESSSERVICECENTER
' STATION
' Only REGULARPICKUP, REQUESTCOURIER, or STATION are
' allowed with international freight shipping.

' -= Packaging =-
sFedExPackaging = "YOURPACKAGING"

' Ground shipping requires "YOURPACKAGING" only.
' Valid values:
' YOURPACKAGING
' FEDEXENVELOPE
' FEDEXPAK
' FEDEXBOX
' FEDEXTUBE
' FEDEX10KGBOX
' FEDEX25KGBOX

' =======End FedEx Services Configuration========
' ===============================================


' ====================USPS=======================
' ===============================================
sUSPSUserID = "xxxxxxxxxxxx"
' ==================END USPS=====================
' ===============================================
%>