<!--#include file="uspsfunctions.asp" -->
<!--#include file="upsfunctions.asp" -->
<!--#include file="fedexfunctions.asp" -->
<%
' Example:
'DisplayAllAsText "10", "50021"

Sub DisplayAllAsSelect(vWeight, sPostalCode, sCountry)
  sUSPSXML = GetUSPSXMLRate(vWeight, sPostalCode)
  If CheckUSPSForErrors(sUSPSXML) Then Exit Sub

  sUPSXML = GetUPSXMLRate(vWeight, sPostalCode, sCountry)
  If CheckUPSForErrors(sUPSXML) Then Exit Sub
  
  'FEDEX ---------------
  sFedExXML = GetFedExXMLRate(vWeight, sPostalCode, sCountry)
  aFedExXML = Split(sFedExXML, "&")
  aTotals = Split(aFedExXML(0), "|")
  aServices = Split(aFedExXML(1), "|")
  'END FEDEX -----------

  'USPS ----------------
  Set oUSPSXML=Server.CreateObject("Microsoft.xmlDOM")
  oUSPSXML.loadxml(sUSPSXML)
  oUSPSXML.getElementsByTagName("Postage")
  'END USPS ------------
  
  'UPS -----------------
  Set oUPSXML=Server.CreateObject("Microsoft.xmlDOM")
  oUPSXML.loadxml(sUPSXML)
  'END UPS -------------

  'Create a select table from the response xml
  %><select name='ASP-Shipping'><%

  'USPS ----------------
  Set oUSPSRates = oUSPSXML.getElementsByTagName("Postage")
  For x = 0 To oUSPSRates.length - 1
    sDisplayString = "USPS " & oUSPSRates.Item(x).selectSingleNode("MailService").Text & " - " & FormatCurrency(Round((oUSPSRates.Item(x).selectSingleNode("Rate").Text * sShippingMarkupFactor + sShippingMarkupFlatRate), 2))
	
    %><option><%=sDisplayString%></option><%
  Next
  'END USPS ------------
  'UPS -----------------
  'Create A Nodelist of All The RatedShipments
  Set NodeList = oUPSXML.documentElement.selectNodes("RatedShipment")
  For x = 0 To NodeList.length - 1
    'Service/Code
    'TotalCharges/MonetaryValue
    sDisplayString = GetFriendlyUPSName(NodeList.Item(x).selectSingleNode("Service/Code").Text) & " - " & FormatCurrency(Round((NodeList.Item(x).selectSingleNode("TotalCharges/MonetaryValue").Text * sShippingMarkupFactor + sShippingMarkupFlatRate), 2))
	
    %><option><%=sDisplayString%></option><%
  Next
 'END UPS -------------
 'FEDEX ---------------
 For x = 0 To Ubound(aServices)
  	If Not IsNumeric(aTotals(x)) Then 
		'aTotals(x) = "0.00 (Err)"
		If sDisplayErrors Then
			%><option><%=aServices(x)%> - <%=aTotals(x)%></option><%
		End If
	Else
		%><option><%=aServices(x)%> - <%=aTotals(x)%></option><%
	End If
  Next
 'END FEDEX -----------  
  %></select><%
End Sub



Sub DisplayAllAsRadio(vWeight, sPostalCode, sCountry)
  sUSPSXML = GetUSPSXMLRate(vWeight, sPostalCode)
  If CheckUSPSForErrors(sUSPSXML) Then Exit Sub

  sUPSXML = GetUPSXMLRate(vWeight, sPostalCode, sCountry)
  If CheckUPSForErrors(sUPSXML) Then Exit Sub
  
  'FEDEX ---------------
  sFedExXML = GetFedExXMLRate(vWeight, sPostalCode, sCountry)
  aFedExXML = Split(sFedExXML, "&")
  aTotals = Split(aFedExXML(0), "|")
  aServices = Split(aFedExXML(1), "|")
  'END FEDEX -----------

  'USPS ----------------
  Set oUSPSXML=Server.CreateObject("Microsoft.xmlDOM")
  oUSPSXML.loadxml(sUSPSXML)
  oUSPSXML.getElementsByTagName("Postage")
  'END USPS ------------
  
  'UPS -----------------
  Set oUPSXML=Server.CreateObject("Microsoft.xmlDOM")
  oUPSXML.loadxml(sUPSXML)
  'END UPS -------------


  'Radio buttons:
  'USPS ----------------
  Set oUSPSRates = oUSPSXML.getElementsByTagName("Postage")
  For x = 0 To oUSPSRates.length - 1
    sDisplayString = "USPS " & oUSPSRates.Item(x).selectSingleNode("MailService").Text & " - " & FormatCurrency(Round((oUSPSRates.Item(x).selectSingleNode("Rate").Text * sShippingMarkupFactor + sShippingMarkupFlatRate), 2))
	
    %><input name="ASP-Shipping" type="radio" value="<%=sDisplayString%>"><%=sDisplayString%><br /><%
  Next
  'END USPS ------------
  'UPS -----------------
  'Create A Nodelist of All The RatedShipments
  Set NodeList = oUPSXML.documentElement.selectNodes("RatedShipment")
  For x = 0 To NodeList.length - 1
    'Service/Code
    'TotalCharges/MonetaryValue
    sDisplayString = GetFriendlyUPSName(NodeList.Item(x).selectSingleNode("Service/Code").Text) & " - " & FormatCurrency(Round((NodeList.Item(x).selectSingleNode("TotalCharges/MonetaryValue").Text * sShippingMarkupFactor + sShippingMarkupFlatRate), 2))
	
    %><input name="ASP-Shipping" type="radio" value="<%=sDisplayString%>"><%=sDisplayString%><br /><%
  Next
 'END UPS -------------
 'FEDEX ---------------
 For x = 0 To Ubound(aServices)
  	If Not IsNumeric(aTotals(x)) Then 
		'aTotals(x) = "0.00 (Err)"
		If sDisplayErrors Then
			%><input name="ASP-Shipping" type="radio" value="<%=aServices(x)%> - <%=aTotals(x)%>"><%=aServices(x)%> - <%=aTotals(x)%><br /><%
		End If
	Else
		%><input name="ASP-Shipping" type="radio" value="<%=aServices(x)%> - <%=aTotals(x)%>"><%=aServices(x)%> - <%=aTotals(x)%><br /><%
	End If
  Next
 'END FEDEX -----------  

End Sub


Sub DisplayAllAsText(vWeight, sPostalCode, sCountry)
  sUSPSXML = GetUSPSXMLRate(vWeight, sPostalCode)
  If CheckUSPSForErrors(sUSPSXML) Then Exit Sub

  sUPSXML = GetUPSXMLRate(vWeight, sPostalCode, sCountry)
  If CheckUPSForErrors(sUPSXML) Then Exit Sub
  
  'FEDEX ---------------
  sFedExXML = GetFedExXMLRate(vWeight, sPostalCode, sCountry)
  aFedExXML = Split(sFedExXML, "&")
  aTotals = Split(aFedExXML(0), "|")
  aServices = Split(aFedExXML(1), "|")
  'END FEDEX -----------

  'USPS ----------------
  Set oUSPSXML=Server.CreateObject("Microsoft.xmlDOM")
  oUSPSXML.loadxml(sUSPSXML)
  oUSPSXML.getElementsByTagName("Postage")
  'END USPS ------------
  
  'UPS -----------------
  Set oUPSXML=Server.CreateObject("Microsoft.xmlDOM")
  oUPSXML.loadxml(sUPSXML)
  'END UPS -------------


  'Plain Text:
  'USPS ----------------
  Set oUSPSRates = oUSPSXML.getElementsByTagName("Postage")
  For x = 0 To oUSPSRates.length - 1
    sDisplayString = "USPS " & oUSPSRates.Item(x).selectSingleNode("MailService").Text & " - " & FormatCurrency(Round((oUSPSRates.Item(x).selectSingleNode("Rate").Text * sShippingMarkupFactor + sShippingMarkupFlatRate), 2))
	
    %><%=sDisplayString%>"<br /><%
  Next
  'END USPS ------------
  'UPS -----------------
  'Create A Nodelist of All The RatedShipments
  Set NodeList = oUPSXML.documentElement.selectNodes("RatedShipment")
  For x = 0 To NodeList.length - 1
    'Service/Code
    'TotalCharges/MonetaryValue
    sDisplayString = GetFriendlyUPSName(NodeList.Item(x).selectSingleNode("Service/Code").Text) & " - " & FormatCurrency(Round((NodeList.Item(x).selectSingleNode("TotalCharges/MonetaryValue").Text * sShippingMarkupFactor + sShippingMarkupFlatRate), 2))
	
    %><%=sDisplayString%><br /><%
  Next
 'END UPS -------------
 'FEDEX ---------------
 For x = 0 To Ubound(aServices)
  	If Not IsNumeric(aTotals(x)) Then 
		'aTotals(x) = "0.00 (Err)"
		If sDisplayErrors Then
			%><%=aServices(x)%> - <%=aTotals(x)%><br /><%
		End If
	Else
		%><%=aServices(x)%> - <%=aTotals(x)%><br /><%
	End If
  Next
 'END FEDEX -----------  

End Sub

%>