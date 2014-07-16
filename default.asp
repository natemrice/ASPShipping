<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="common.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>ASPShipping FOSS (Free Open Source Software) 04/15/2007</title>
<style type="text/css">
<!--
.style1 {
	color: #FFFFFF;
	font-family: Geneva, Arial, Helvetica, sans-serif;
}
.style2 {font-family: Georgia, "Times New Roman", Times, serif}
.style3 {color: #FFFFFF; font-family: Georgia, "Times New Roman", Times, serif; }
.ImageFloatRight {float:right}
-->
</style>
</head>

<body>
<table width="90%" border="0" align="center" cellpadding="5" cellspacing="0">
  <tr bgcolor="#3D81EE">
  	<td colspan="3" align="center"><script type="text/javascript"><!--
google_ad_client = "pub-8923326944333262";
google_ad_width = 728;
google_ad_height = 90;
google_ad_format = "728x90_as";
google_ad_type = "text_image";
//2007-04-16: ASPShipping
google_ad_channel = "5122933448";
google_color_border = "3D81EE";
google_color_bg = "FFFFFF";
google_color_link = "0000FF";
google_color_text = "000000";
google_color_url = "008000";
//-->
</script>
<script type="text/javascript"
  src="http://pagead2.googlesyndication.com/pagead/show_ads.js">
</script></td>
  </tr>
  <tr>
    <td bgcolor="#3D81EE">&nbsp;</td>
    <td><h2><a href="http://www.naterice.com/">NateRice.com</a> ASPShipping FOSS (Free Open Source Software) <img src="../images/geek.gif" width="19" height="19" /></h2>
      <h3>Current Rev: 05/18/2007</h3>
              Past Rev: 04/17/2007
              <p><a href="http://www.aspin.com/func/review?tree=aspin/tutorial/xml&amp;id=8430610"><img src="../images/aspin1002.gif" alt="Rate this script on ASPIN.COM" border="0" class="ImageFloatRight" /></a><a href="NateRice.com_ASPShipping_04-15-2007.zip"><img src="http://www.naterice.com/blog/images/3/winzip.gif" alt="Download ASPShipping" width="20" height="20" border="0" /></a> <a href="NateRice.com_ASPShipping_05-18-2007.zip">Download Source Code</a> | <a href="changelog.html">Change Log / Liscense</a></p>
      <p>Since there are no free open source ASP solutions for live shipping rates with UPS, FedEx and USPS I decided that I'm going to write a free open source solution that anyone can use. This code is all 100% ASP and requires no non-Microsoft plugins or COM+ dlls. </p>
      <p>Right now USPS, UPS and FedEx now have working code. And I will continue to improve them over the coming weeks and months. </p>
      <p>If you like this code please consider posting a note on your favorite forum, a link on your personal website or maybe even a bookmark from del.icio.us. Spreading the word is one of the most important ways of &quot;giving back&quot;. </p>
      <p>Keep checking for updates as I will release them in a date stamped manner. Future updates will include enabling and disabling of all types of services from all service providers, enabling and disabling of specific providers, and eventually I'll probably be releasing a VB.NET version of this whole suite.</p>
      <p>Any and all feedback is appreciated so please feel free to email me at the address listed in this image:</p>
      <p><img src="http://www.naterice.com/blog/images/1/contact.png" width="173" height="34" /> </p>
      <p>Integration is simple, just add an &quot;include&quot; statement to the top of an  existing page, and then get the rates with a few simple lines of code.  Integration can take place in as little as 2 lines of ASP code.</p>
      <p>You can currently render the XML response as either a select box, an array of radio buttons, or just plain old text if you'd just like to display it and not use it in a form.</p>
      <p>ASPShipping Example:</p>
      <form id="shipping" name="shipping" method="post" action="">
        <p>Postal Code:
          <input name="postalcode" type="text" id="postalcode" value="50021" maxlength="10" />
        </p>
        <p>2 Letter Country Code:
          <input name="country" type="text" id="country" value="US" />
        </p>
        <p>Weight (lbs):
          <input name="weight" type="text" id="weight" value="10" />
        </p>
        <p>
          <input type="submit" name="Submit" value="Submit" />
        </p>
      </form>
      <%
If len(Request.Form("submit")) > 0 Then
	sTestUPSXML = GetUPSXMLRate(Request.Form("weight"), Request.Form("postalcode"), Request.Form("country"))

	%>
      <h2>LIVE UPS, USPS and FedEx Shipping Rate Estimates</h2>
      Shipping To: <%=Request.Form("postalcode")%>/<%=Request.Form("country")%><br />
Weight: <%=Request.Form("weight")%><br/>
<h3>Configuration:</h3>
Display Errors: <%=sDisplayErrors%><br/>
Shipper Postal Code: <%=sShipperPostalCode%><br/>

Markup Factor: <%=sShippingMarkupFactor%><br/>
Markup Flat Rate: <%=sShippingMarkupFlatRate%><br/>
<h3>Select Box:</h3>
<%
	DisplayAllAsSelect Request.Form("weight"), Request.Form("postalcode"), Request.Form("country")
	%>
<h3>Radio Buttons:</h3>
<%
	DisplayAllAsRadio Request.Form("weight"), Request.Form("postalcode"), Request.Form("country")
	%>
<h3>Plain Text:</h3>
<%
	DisplayAllAsText Request.Form("weight"), Request.Form("postalcode"), Request.Form("country")

End If
%>
<p></p></td>
    <td bgcolor="#3D81EE"><br /><br /><script type="text/javascript"><!--
google_ad_client = "pub-8923326944333262";
google_ad_width = 160;
google_ad_height = 600;
google_ad_format = "160x600_as";
google_ad_type = "text_image";
//2007-04-16: ASPShipping
google_ad_channel = "5122933448";
google_color_border = "3D81EE";
google_color_bg = "FFFFFF";
google_color_link = "0000FF";
google_color_text = "000000";
google_color_url = "008000";
//-->
</script>
<script type="text/javascript"
  src="http://pagead2.googlesyndication.com/pagead/show_ads.js">
</script></td>
  </tr>
  <tr bgcolor="#3D81EE">
  	<td colspan="3">&nbsp;</td>
  </tr>
</table>
<h2>&nbsp;</h2>
</body>
</html>
