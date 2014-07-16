<!--#include file="configuration.asp" -->

<%
'sUSPSRateXML = GetUSPSXMLRate("20", "50021")

'DisplayUSPSXMLRateAsSelect(sUSPSRateXML)


Sub DisplayUSPSXMLRateAsSelect(sUSPSXML)
  'Here we will check for hard errors from USPS
  If CheckUSPSForErrors(sUSPSXML) Then Exit Sub

  Set oUSPSXML=Server.CreateObject("Microsoft.xmlDOM")
  oUSPSXML.loadxml(sUSPSXML)
  
  oUSPSXML.getElementsByTagName("Postage")
  
  'Create a select table from the response xml
  %><select name='USPS-Shipping'><%

  Set oUSPSRates = oUSPSXML.getElementsByTagName("Postage")
  For x = 0 To oUSPSRates.length - 1
    sDisplayString = "USPS " & oUSPSRates.Item(x).selectSingleNode("MailService").Text & " - " & FormatCurrency(Round((oUSPSRates.Item(x).selectSingleNode("Rate").Text * sShippingMarkupFactor + sShippingMarkupFlatRate), 2))
	
    %><option><%=sDisplayString%></option><%
  Next
  %></select><%
End Sub

Sub DisplayUSPSXMLRateAsRadio(sUSPSXML)
  'Here we will check for hard errors from USPS
  If CheckUSPSForErrors(sUSPSXML) Then Exit Sub

  Set oUSPSXML=Server.CreateObject("Microsoft.xmlDOM")
  oUSPSXML.loadxml(sUSPSXML)
  
  oUSPSXML.getElementsByTagName("Postage")

  Set oUSPSRates = oUSPSXML.getElementsByTagName("Postage")
  For x = 0 To oUSPSRates.length - 1
    sDisplayString = "USPS " & oUSPSRates.Item(x).selectSingleNode("MailService").Text & " - " & FormatCurrency(Round((oUSPSRates.Item(x).selectSingleNode("Rate").Text * sShippingMarkupFactor + sShippingMarkupFlatRate), 2))

    %><input name="USPS-Shipping" type="radio" value="<%=sDisplayString%>"><%=sDisplayString%><br /><%
  Next
End Sub

Sub DisplayUSPSXMLRateAsText(sUSPSXML)
  'Here we will check for hard errors from USPS
  If CheckUSPSForErrors(sUSPSXML) Then Exit Sub

  Set oUSPSXML=Server.CreateObject("Microsoft.xmlDOM")
  oUSPSXML.loadxml(sUSPSXML)
  
  oUSPSXML.getElementsByTagName("Postage")

  Set oUSPSRates = oUSPSXML.getElementsByTagName("Postage")
  For x = 0 To oUSPSRates.length - 1
    sDisplayString = "USPS " & oUSPSRates.Item(x).selectSingleNode("MailService").Text & " - " & FormatCurrency(Round((oUSPSRates.Item(x).selectSingleNode("Rate").Text * sShippingMarkupFactor + sShippingMarkupFlatRate), 2))

    %><%=sDisplayString%><br /><%
  Next
End Sub

Function GetUSPSXMLRate(vTotalWeight, sDestinationPostalCode)
  sUSPSXML = BuildUSPSXML(vTotalWeight, sDestinationPostalCode)

  Set oXMLHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP")
  ' Production Server = http://production.shippingapis.com/ShippingAPI.dll?
  ' Test Server = http://testing.shippingapis.com/ShippingAPITest.dll?
  oXMLHTTP.Open "GET","http://production.shippingapis.com/ShippingAPI.dll?" & sUSPSXML,false
  oXMLHTTP.setRequestHeader "Content-Type", "text/xml"
  
  On Error Resume Next
  oXMLHTTP.send ""
  
  If Err.Number <> 0 Then
    GetUSPSXMLRate = "Error retrieving USPS quote. (site unavailable)"
	Exit Function
  End If
  On Error Goto 0
  
  GetUSPSXMLRate = oXMLHTTP.responseText
End Function


' This is the PRODUCTION BuildUSPSXML function
' rename it to enable. Rename the test function
' if you rename this to disable test.
Function BuildUSPSXML(sWeight, sDestinationPostalCode)

  sXML = sXML & "API=RateV2&XML="
  sXML = sXML & "<RateV2Request USERID=""" & sUSPSUserID & """>"
  sXML = sXML & "	<Package ID=""0"">"
  sXML = sXML & "		<Service>All</Service>"
  sXML = sXML & "		<ZipOrigination>" & sShipperPostalCode & "</ZipOrigination>"
  sXML = sXML & "		<ZipDestination>" & sDestinationPostalCode & "</ZipDestination>"
  sXML = sXML & "		<Pounds>" & sWeight & "</Pounds>"
  sXML = sXML & "		<Ounces>0</Ounces>"
  sXML = sXML & "		<Container>Flat Rate Box</Container>"
  sXML = sXML & "		<Size>REGULAR</Size>"
  sXML = sXML & "		<Machinable>False</Machinable>"
  sXML = sXML & "	</Package>"
  sXML = sXML & "</RateV2Request>"
  
  BuildUSPSXML = Replace(sXML, vbTab, "")
End Function
Function BuildUSPSXMLTesting(sWeight, sDestinationPostalCode)

  sXML = sXML & "API=RateV2&XML="
  sXML = sXML & "<RateV2Request USERID=""" & sUSPSUserID & """>"
  sXML = sXML & "	<Package ID=""0"">"
  sXML = sXML & "		<Service>PRIORITY</Service>"
  sXML = sXML & "		<ZipOrigination>10022</ZipOrigination>"
  sXML = sXML & "		<ZipDestination>20008</ZipDestination>"
  sXML = sXML & "		<Pounds>10</Pounds>"
  sXML = sXML & "		<Ounces>5</Ounces>"
  sXML = sXML & "		<Container>Flat Rate Box</Container>"
  sXML = sXML & "		<Size>REGULAR</Size>"
  sXML = sXML & "	</Package>"
  sXML = sXML & "</RateV2Request>"
  
  BuildUSPSXML = Replace(sXML, vbTab, "")
End Function
' ============================================================



Function CheckUSPSForErrors(sUSPSResponseXML)
  If InStr(sUSPSResponseXML, "Error retrieving USPS quote") > 0 Then
    CheckUSPSForErrors = True
	Exit Function
  End IF

  Set oUSPSXML=Server.CreateObject("Microsoft.xmlDOM")
  oUSPSXML.loadxml(sUSPSResponseXML)

  Set oError = oUSPSXML.getElementsByTagName("Error")

  If sDisplayErrors AND oError.length > 0 Then
    sErrorMessage = oError.Item(0).selectSingleNode("Number").Text & " - " & oError.Item(0).selectSingleNode("Description").Text
	%><%=sErrorMessage%><br /><%
	CheckUSPSForErrors = True
  Else
    CheckUSPSForErrors = False
  End If

End Function
%>