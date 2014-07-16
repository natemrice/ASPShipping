<!--#include file="configuration.asp" -->

<%
Sub DisplayUPSXMLRateAsSelect(sUPSXML)
  'Here we will check for hard errors from UPS
  If CheckUPSForErrors(sUPSXML) Then Exit Sub

  Set oUPSXML=Server.CreateObject("Microsoft.xmlDOM")
  oUPSXML.loadxml(sUPSXML)
  
  'Create a select table from the response xml
  %>
<select name='UPS-Shipping'><%

  'Create A Nodelist of All The RatedShipments
  Set NodeList = oUPSXML.documentElement.selectNodes("RatedShipment")
  For x = 0 To NodeList.length - 1
    'Service/Code
    'TotalCharges/MonetaryValue
    sDisplayString = GetFriendlyUPSName(NodeList.Item(x).selectSingleNode("Service/Code").Text) & " - " & FormatCurrency(Round((NodeList.Item(x).selectSingleNode("TotalCharges/MonetaryValue").Text * sShippingMarkupFactor + sShippingMarkupFlatRate), 2))
	
    %><option><%=sDisplayString%></option><%
  Next
  %></select><%
End Sub

Sub DisplayUPSXMLRateAsRadio(sUPSXML)
  'Here we will check for hard errors from UPS
  If CheckUPSForErrors(sUPSXML) Then Exit Sub

  Set oUPSXML=Server.CreateObject("Microsoft.xmlDOM")
  oUPSXML.loadxml(sUPSXML)

  'Create A Nodelist of All The RatedShipments
  Set NodeList = oUPSXML.documentElement.selectNodes("RatedShipment")
  For x = 0 To NodeList.length - 1
    'Service/Code
    'TotalCharges/MonetaryValue
    sDisplayString = GetFriendlyUPSName(NodeList.Item(x).selectSingleNode("Service/Code").Text) & " - " & FormatCurrency(Round((NodeList.Item(x).selectSingleNode("TotalCharges/MonetaryValue").Text * sShippingMarkupFactor + sShippingMarkupFlatRate), 2))

    %><input name="UPS-Shipping" type="radio" value="<%=sDisplayString%>"><%=sDisplayString%><br /><%
  Next
End Sub

Sub DisplayUPSXMLRateAsText(sUPSXML)
  'Here we will check for hard errors from UPS
  If CheckUPSForErrors(sUPSXML) Then Exit Sub

  Set oUPSXML=Server.CreateObject("Microsoft.xmlDOM")
  oUPSXML.loadxml(sUPSXML)

  'Create A Nodelist of All The RatedShipments
  Set NodeList = oUPSXML.documentElement.selectNodes("RatedShipment")
  For x = 0 To NodeList.length - 1
    'Service/Code
    'TotalCharges/MonetaryValue
    sDisplayString = GetFriendlyUPSName(NodeList.Item(x).selectSingleNode("Service/Code").Text) & " - " & FormatCurrency(Round((NodeList.Item(x).selectSingleNode("TotalCharges/MonetaryValue").Text * sShippingMarkupFactor + sShippingMarkupFlatRate), 2))

    %><%=sDisplayString%><br /><%
  Next
End Sub

Function GetUPSXMLRate(vTotalWeight, sDestinationPostalCode, sDestinationCountryCode)
  Set oXMLHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP")
  oXMLHTTP.Open "POST","https://www.ups.com/ups.app/xml/Rate?",false
  oXMLHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
  
  sUPSXML = BuildUPSXML(vTotalWeight, sDestinationPostalCode, sDestinationCountryCode)
  
  ' Now we are doing the actual post of XML to the UPS Servers.
  ' There is a chance this will fail if for some reason we can't
  ' reach the UPS servers.
  On Error Resume Next
  oXMLHTTP.send sUPSXML
  
  If Err.Number <> 0 Then
    GetUPSXMLRate = "Error retrieving UPS quote. (site unavailable)"
	Exit Function
  End If
  On Error Goto 0
  
  GetUPSXMLRate = oXMLHTTP.responseText
End Function


Function BuildUPSXML(sWeight, sDestinationPostalCode, sDestinationCountryCode)

  sXML = sXML & "<?xml version='1.0'?>"
  sXML = sXML & "	<AccessRequest xml:lang='en-US'>"
  sXML = sXML & "		<AccessLicenseNumber>" & sUPSAccessLiscenseNumber & "</AccessLicenseNumber>"
  sXML = sXML & "		<UserId>" & sUPSUserID & "</UserId>"
  sXML = sXML & "		<Password>" & sUPSPassword & "</Password>"
  sXML = sXML & "	</AccessRequest>"
  sXML = sXML & "<?xml version='1.0'?>"
  sXML = sXML & "	<RatingServiceSelectionRequest xml:lang='en-US'>"
  sXML = sXML & "		<Request>"
  sXML = sXML & "			<TransactionReference>"
  sXML = sXML & "				<CustomerContext>Rating and Service</CustomerContext>"
  sXML = sXML & "				<XpciVersion>1.0001</XpciVersion>"
  sXML = sXML & "			</TransactionReference>"
  sXML = sXML & "			<RequestAction>Rate</RequestAction>"
  sXML = sXML & "			<RequestOption>shop</RequestOption>"
  sXML = sXML & "		</Request>"
  sXML = sXML & "		<PickupType>"
  sXML = sXML & "			<Code>01</Code>"
  sXML = sXML & "		</PickupType>"
  sXML = sXML & "		<Shipment>"
  sXML = sXML & "			<Shipper>"
  sXML = sXML & "				<Address>"
  sXML = sXML & "					<PostalCode>" & sShipperPostalCode & "</PostalCode>"
  sXML = sXML & "				</Address>"
  sXML = sXML & "			</Shipper>"
  sXML = sXML & "			<ShipTo>"
  sXML = sXML & "				<Address>"
  sXML = sXML & "					<PostalCode>" & sDestinationPostalCode & "</PostalCode>"
  sXML = sXML & "					<CountryCode>" & sDestinationCountryCode & "</CountryCode>"
  sXML = sXML & "				</Address>"
  sXML = sXML & "			</ShipTo>"
  sXML = sXML & "			<Service>"
  sXML = sXML & "				<Code>11</Code>"
  sXML = sXML & "			</Service>"
  sXML = sXML & "			<Package>"
  sXML = sXML & "				<PackagingType>"
  sXML = sXML & "					<Code>02</Code>"
  sXML = sXML & "					<Description>Package</Description>"
  sXML = sXML & "				</PackagingType>"
  sXML = sXML & "				<Description>Rate Shopping</Description>"
  sXML = sXML & "				<PackageWeight>"
  sXML = sXML & "					<Weight>" & sWeight & "</Weight>"
  sXML = sXML & "				</PackageWeight>"
  sXML = sXML & "			</Package>"
  sXML = sXML & "			<ShipmentServiceOptions/>"
  sXML = sXML & "		</Shipment>"
  sXML = sXML & "</RatingServiceSelectionRequest>"
  
  BuildUPSXML = Replace(sXML, vbTab, "")
End Function

Function GetFriendlyUPSName(vCode)
  Select Case vCode
    Case "01"
      GetFriendlyUPSName = "UPS Next Day Air"
    Case "02"
      GetFriendlyUPSName = "UPS 2nd Day Air"
    Case "03"
      GetFriendlyUPSName = "UPS Ground"
    Case "07"
      GetFriendlyUPSName = "UPS Worldwide Express"
    Case "08"
      GetFriendlyUPSName = "UPS Worldwide Expedited"
    Case "11"
      GetFriendlyUPSName = "UPS Standard"
    Case "12"
      GetFriendlyUPSName = "UPS 3 Day Select"
    Case "13"
      GetFriendlyUPSName = "UPS Next Day Air Saver"
    Case "14"
      GetFriendlyUPSName = "UPS Next Day Air Early A.M."
    Case "54"
      GetFriendlyUPSName = "UPS Worldwide Express Plus"
    Case "59"
      GetFriendlyUPSName = "UPS 2nd Day Air A.M."
    Case "65"
      GetFriendlyUPSName = "UPS Saver"
  End Select
End Function

Function CheckUPSForErrors(sUPSResponseXML)
  If InStr(sUPSResponseXML, "Error retrieving UPS quote") > 0 Then
    CheckUPSForErrors = True
	Exit Function
  End IF

  Set oUPSXML=Server.CreateObject("Microsoft.xmlDOM")
  oUPSXML.loadxml(sUPSResponseXML)
  
  Set NodeList = oUPSXML.documentElement.selectNodes("Response")
  sStatus = NodeList.Item(0).selectSingleNode("ResponseStatusCode").Text
  'Set NodeList = Nothing

  If sStatus = 0 Then
	'An error occured
    If sDisplayErrors Then
	  Set oErrorNodeList = oUPSXML.getElementsByTagName("Error")
      sErrorMessage = oErrorNodeList.Item(0).selectSingleNode("ErrorCode").Text & " - " & oErrorNodeList.Item(0).selectSingleNode("ErrorDescription").Text
	  %><br /><%=sErrorMessage%><%
    End If
    CheckUPSForErrors = True
  Else
    CheckUPSForErrors = False
  End If
End Function
%>