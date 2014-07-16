<!--#include file="configuration.asp" -->
<%
' FedEx Functions file. FedEx is very different from UPS and USPS in the way that they handle XML. Follow the steps
' below to successfully retrieve rate quotes from FedEx.
' Step 1) Update the "configuration.asp" file:
'           Change any options relevent to your configuration. Update the "sFEDEXAccountNumber" string with your FedEx
'           account number. You won't have a meter number yet so simply leave the value blank. Pay special attention
'           to the "Meter Number" section of the configuration file. You'll have to send your personal information
'           to successfully retrieve a meter number. Once you retrieve a meter number the values in the "Meter Number"
'           section of the configuration file will never be used again.
' Step 2) Uncomment the "DisplayMeterNumber" function and view this page in a browser:
'           Uncomment the line below and load this page in a browser to retrieve a FedEx meter number. This must be
'           done for each machine you intend to run this code on. Once the code successfully returns a meter number
'           you'll have to copy and paste it into "configuration.asp" file and comment out this line once again.
'
'vvvvvvvv Uncomment Line Below To Display Meter Number vvvvvvvvvvv
'DisplayMeterNumber(GetFedExMeterNumberXML)
'^^^^^^^^ Uncomment Line Above To Display Meter Number ^^^^^^^^^^^
'
' Step 3) Copy Meter Number to "configuration.asp" file and re-comment out the line above.
'           Once we successfully have the meter number we can run any FedEx subs such as "DisplayFedExXMLAsSelect",
'           "DisplayFedExXMLAsRadio", or "DisplayFedExXMLAsText". You can also run any of the "common" subs such as
'           "DisplayAllAsSelect", "DisplayAllAsRadio", or "DisplayAllAsText".
' Step 4) Integration of some or all rate services.
'			If you wish to call a common function you simply have to "include" the "common.asp" file and call functions
'           normally. If you want to include only one rate service you can optionally "include" only the service you
'           are interested in using such as "fedexfunctions.asp" or "upsfunctions.asp"
'           
'             Example Include: <!--#include file="uspsfunctions.asp" -->

Sub DisplayFedExXMLAsSelect(sTotals, sServices)
  aTotals = Split(sTotals, "|")
  aServices = Split(sServices, "|")

  %><select name='FedEx-Shipping'><%
  For x = 0 To Ubound(aServices)
  	If Not IsNumeric(aTotals(x)) Then 
		aTotals(x) = "0.00 (Err)"
		If sDisplayErrors Then
			%><option><%=aServices(x)%> - <%=aTotals(x)%></option><%
		End If
	Else
		%><option><%=aServices(x)%> - <%=aTotals(x)%></option><%
	End If
  Next
  %></select><%
End Sub

Sub DisplayFedExXMLAsRadio(sTotals, sServices)
  aTotals = Split(sTotals, "|")
  aServices = Split(sServices, "|")

  For x = 0 To Ubound(aServices)
  	If Not IsNumeric(aTotals(x)) Then 
		aTotals(x) = "0.00 (Err)"
		If sDisplayErrors Then
			%><input name="FedEx-Shipping" type="radio" value="<%=aServices(x)%> - <%=aTotals(x)%>"><%=aServices(x)%> - <%=aTotals(x)%><br /><%
		End If
	Else
		%><input name="FedEx-Shipping" type="radio" value="<%=aServices(x)%> - <%=aTotals(x)%>"><%=aServices(x)%> - <%=aTotals(x)%><br /><%
	End If
  Next
End Sub

Sub DisplayFedExXMLAsText(sTotals, sServices)
  aTotals = Split(sTotals, "|")
  aServices = Split(sServices, "|")

  For x = 0 To Ubound(aServices)
  	If Not IsNumeric(aTotals(x)) Then 
		aTotals(x) = "0.00 (Err)"
		If sDisplayErrors Then
			%><%=aServices(x)%> - <%=aTotals(x)%><%
		End If
	Else
		%><%=aServices(x)%> - <%=aTotals(x)%><%
	End If
  Next
End Sub

Function DisplayMeterNumber(sXML)
	'This function is used to retrieve a meter number from FedEx. Once a meter number has been successfully
	'retrieved this function is never used again.
	If InStr(sXML, "<MeterNumber>") > 0 Then
		aMeterNumber1 = Split(sXML, "<MeterNumber>")
		aMeterNumber2 = Split(aMeterNumber1(1), "</MeterNumber>")
		
		sMeterNumber = aMeterNumber2(0)
		
		sMeterInfo = "<strong>Meter Number</strong>: " & sMeterNumber & "<br />" & _
		"(you must copy and paste this number into the<br />" & _
		"configuration.asp file under the variable<br />" & _
		"named ""sFEDEXMeterNumber"". This is a one time<br />" & _
		"event for each new server that retrieves<br />" & _
		"FedEx quotes.)"
		
		DisplayMeterNumber = sMeterInfo
	Else
		DisplayMeterNumber = sXML
	End If
End Function

Function GetFedExMeterNumberXML
  'This function is used to retrieve a FedEx meter number. You'll need a meter number to retrieve rate quotes from
  'FedEx. The idea being that each machine you run code on will have a different meter number.

  Set oXMLHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP")
  oXMLHTTP.Open "POST","https://gateway.fedex.com/GatewayDC",false
  oXMLHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
  
  sFedExMeterNumberXML = BuildMeterNumberXML
  
  ' Now we are doing the actual post of XML to the FedEx Servers.
  ' There is a chance this will fail if for some reason we can't
  ' reach the FedEx servers.
  On Error Resume Next
  oXMLHTTP.send sFedExMeterNumberXML
  
  If Err.Number <> 0 Then
    GetFedExMeterNumberXML = "Error retrieving meter number. (site unavailable)"
	Exit Function
  End If
  On Error Goto 0
  
  sXMLResponse = oXMLHTTP.responseText
  
  GetFedExMeterNumberXML = DisplayMeterNumber(sXMLResponse)
End Function

Function GetFedExXMLRate(sWeight, sDestinationPostalCode, sDestinationCountryCode)
	'This function is used to do the rate requests from FedEx. This function will
	'do multiple XML posts to FedEx. Requesting rates for each service enabled in
	'the common.asp file under "FedEx Services Configuration" section.
	
	aEnabledFedExServices = Split(sEnabledFedExServices, ",")

	For Each sFedExService In aEnabledFedExServices
		sFedExService = Trim(sFedExService)
		
		If InStr(LCase(sFedExService), "international") > 0 And UCase(sDestinationCountryCode) = "US" Then
			'We don't want to query a international service if the destination country code is domestic.
			'There is no point and it will always return an error.
			'skip
		ElseIf InStr(LCase(sFedExService), "international") = 0 And UCase(sDestinationCountryCode) <> "US" Then
			'We also don't want to query a domestic shipping service if the destination country code is
			'not "US". So skip this too.
			'skip
		Else
			'Here we will do the query for the FedEx service we want to retrieve a rate from. First
			'thing we have to do is build a new XML request for the specific service we are requesting
			'a rate from and then submit that XML to FedEx.
			
			'bulding XML
			sFedExRateXML = BuildFedExRateXML(sWeight, sDestinationPostalCode, sDestinationCountryCode, sFedExService)
		
			Set oXMLHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP")
  			oXMLHTTP.Open "POST","https://gateway.fedex.com/GatewayDC",false
  			oXMLHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"

	        On Error Resume Next
			'submit XML to fedex
    	    oXMLHTTP.send sFedExRateXML
  
  			If Err.Number <> 0 Then
				'In some instances the send method of MSXML2 will return an error if the FedEx site is
				'unavailable for some reason. If so we will handle the error instead of displaying an
				'asp error page. There is a possiblity that we can do more error handling here at some
				'point in the future. If someone wants to submit an error handling function that'd be
				'great.
   				GetFEDEXXMLRate = "Error retrieving sFEDEX quote. (site unavailable)"
				Exit Function
  			End If
  			On Error Goto 0
		
			'here is FedEx's response in a string that is in XML format.
			sXMLResponse = oXMLHTTP.responseText
		
			If InStr(sXMLResponse, "<NetCharge>") > 0 Then
				'we successfully returned a rate from fedex. now we
				'parse it from the XML and store the value.
				aTotal0 = Split(sXMLResponse, "<NetCharge>")
				aTotal1 = Split(aTotal0(1), "</NetCharge>")
				sTotal = aTotal1(0)
			Else
				'we recieved some other response from FedEx (probably
				'an error). Store the error.
				sTotal = sXMLResponse
			End If

			'Now we store the total we retrieved from FedEx in a pipe ("|") delimited string "sTotals",
			'along with the service name in another pipe delimited string "sServices". We can quickly
			'parse this with other functions.
			If len(sTotals) = 0 Then
				sTotals = sTotal
				sServices = GetFriendlyFedExName(sFedExService)
			Else
				sTotals = sTotals & "|" & sTotal
				sServices = sServices & "|" & GetFriendlyFedExName(sFedExService)
  			End If
		End If
	Next
	
	'Now return rates and services and seperate the two with pipe delimited strings with a "&".
	GetFedExXMLRate = sTotals & "&" & sServices
End Function

Function BuildMeterNumberXML
	'This function is only used when building the XML used to submit a request
	'for a meter number from FedEx. After successfully retrieving a meter number
	'it's never used again.
	sXML = sXML & "<?xml version=""1.0"" encoding=""UTF-8"" ?>"
	sXML = sXML & "<FDXSubscriptionRequest  xmlns:api=""http://www.fedex.com/fsmapi"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:noNamespaceSchemaLocation=""FDXSubscriptionRequest.xsd"">"
	sXML = sXML & "	<RequestHeader>"
	sXML = sXML & "		<CustomerTransactionIdentifier>String</CustomerTransactionIdentifier>"
	sXML = sXML & "		<AccountNumber>" & sFEDEXAccountNumber & "</AccountNumber>"
	sXML = sXML & "	</RequestHeader>"
	sXML = sXML & "	<Contact>"
	sXML = sXML & "		<PersonName>" & sFEDEXName &"</PersonName>"
	sXML = sXML & "		<CompanyName>" & sFEDEXCompanyName & "</CompanyName>"
	sXML = sXML & "		<Department>Shipping</Department>"
	sXML = sXML & "		<PhoneNumber>" & sFEDEXPhoneNumber & "</PhoneNumber>"
	sXML = sXML & "		<E-MailAddress>" & sFEDEXEmailAddress & "</E-MailAddress>"
	sXML = sXML & "	</Contact>"
	sXML = sXML & "	<Address>"
	sXML = sXML & "		<Line1>" & sFEDEXAddress & "</Line1>"
	sXML = sXML & "		<City>" & sFEDEXCity & "</City>"
	sXML = sXML & "		<StateOrProvinceCode>" & sFEDEXState & "</StateOrProvinceCode>"
	sXML = sXML & "		<PostalCode>" & sShipperPostalCode & "</PostalCode>"
	sXML = sXML & "		<CountryCode>US</CountryCode>"
	sXML = sXML & "	</Address>"
	sXML = sXML & "</FDXSubscriptionRequest>"

	BuildMeterNumberXML = Replace(sXML, vbTab, "")
End Function

Function BuildFedExRateXML(sWeight, sDestinationPostalCode, sDestinationCountryCode, sServiceType)
	'This function builds a standard FedEx rate request XML.
	
	'For some odd reason FedEx also wants at least one decimal place for numeric weight values.
	'So if this is a whole number append a ".0" so as to appease the FexEx gods.
	If InStr(sWeight, ".") > 0 Then
	  'no append
	Else
	  sWeight = sWeight & ".0"
	End If

	'FedEx is apparently devided into 2 services. One "ground" and one "express". Services that have
	'"ground" in the name are ground services, all others fall under "express" services.
	If InStr(LCase(sServiceType), "ground") > 0 Then
		sFedExCompany = "FDXG"
	Else
		sFedExCompany = "FDXE"
	End If

	sXML = sXML & "<?xml version=""1.0"" encoding=""UTF-8"" ?>"
	sXML = sXML & "<FDXRateRequest xmlns:api=""http://www.fedex.com/fsmapi"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:noNamespaceSchemaLocation=""FDXRateRequest.xsd"">"
	sXML = sXML & "		<RequestHeader>"
	sXML = sXML & "			<CustomerTransactionIdentifier>CTIString</CustomerTransactionIdentifier>"
	sXML = sXML & "			<AccountNumber>" & sFEDEXAccountNumber & "</AccountNumber>"
	sXML = sXML & "			<MeterNumber>" & sFEDEXMeterNumber & "</MeterNumber>"
	sXML = sXML & "			<CarrierCode>" & sFedExCompany & "</CarrierCode>"
	sXML = sXML & "		</RequestHeader>"
	sXML = sXML & "		<DropoffType>" & sFedExDropOffType & "</DropoffType>"
	sXML = sXML & "		<Service>" & sServiceType & "</Service>"
	sXML = sXML & "		<Packaging>" & sFedExPackaging & "</Packaging>"
	sXML = sXML & "		<WeightUnits>LBS</WeightUnits>"
	sXML = sXML & "		<Weight>" & sWeight & "</Weight>"
	sXML = sXML & "		<OriginAddress>"
	sXML = sXML & "			<PostalCode>" & sShipperPostalCode & "</PostalCode>"
	sXML = sXML & "			<CountryCode>" & sShipperCountryCode & "</CountryCode>"
	sXML = sXML & "		</OriginAddress>"
	sXML = sXML & "		<DestinationAddress>"
	sXML = sXML & "			<PostalCode>" & sDestinationPostalCode & "</PostalCode>"
	sXML = sXML & "			<CountryCode>" & sDestinationCountryCode & "</CountryCode>"
	sXML = sXML & "		</DestinationAddress>"
	sXML = sXML & "		<Payment>"
	sXML = sXML & "			<PayorType>SENDER</PayorType>"
	sXML = sXML & "		</Payment>"
	sXML = sXML & "		<PackageCount>1</PackageCount>"
	sXML = sXML & "</FDXRateRequest>"

	BuildFedExRateXML = Replace(sXML, vbTab, "")
End Function




Function GetFriendlyFedExName(sCode)
  'This function just resolves the FedEx XML standard service name to a friendly name for people to read.
  Select Case sCode
    Case "PRIORITYOVERNIGHT"
      GetFriendlyFedExName = "FedEx Priority Overnight"
    Case "STANDARDOVERNIGHT"
      GetFriendlyFedExName = "FedEx Standard Overnight"
    Case "FIRSTOVERNIGHT"
      GetFriendlyFedExName = "FedEx First Overnight"
    Case "FEDEX2DAY"
      GetFriendlyFedExName = "FedEx 2 Day"
    Case "FEDEXEXPRESSSAVER"
      GetFriendlyFedExName = "FedEx Express Saver"
    Case "INTERNATIONALPRIORITY"
      GetFriendlyFedExName = "FedEx International Priority"
    Case "INTERNATIONALECONOMY"
      GetFriendlyFedExName = "FedEx International Economy"
    Case "INTERNATIONALFIRST"
      GetFriendlyFedExName = "FedEx International First"
    Case "FEDEX1DAYFREIGHT"
      GetFriendlyFedExName = "FedEx 1 Day Freight"
    Case "FEDEX2DAYFREIGHT"
      GetFriendlyFedExName = "FedEx 2 Day Freight"
    Case "FEDEX3DAYFREIGHT"
      GetFriendlyFedExName = "FedEx 3 Day Freight"
    Case "FEDEXGROUND"
      GetFriendlyFedExName = "FedEx Ground"
    Case "GROUNDHOMEDELIVERY"
      GetFriendlyFedExName = "FedEx Ground Home Delivery"
    Case "INTERNATIONALPRIORITY FREIGHT"
      GetFriendlyFedExName = "FedEx Int. Priority Freight"
    Case "INTERNATIONALECONOMY FREIGHT"
      GetFriendlyFedExName = "FedEx Int. Economy Freight"
	Case "EUROPEFIRSTINTERNATIONALPRIORITY"
	  GetFriendlyFedExName = "FedEx Euro. First Int. Priority"
  End Select
End Function
%>