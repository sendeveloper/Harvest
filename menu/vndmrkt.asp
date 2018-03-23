<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<%
     Session("redirect") = ""

     ThisPage = "https://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL")
%>

<!--#include virtual="includes/connection.asp"-->

	<head>
		<meta http-equiv="content-type" content="text/html; charset=utf-8" />
		<title>Harvest American, Inc. | Vendors & Marketing</title>
		<link href="<%=strPathMenuIncludes%>ha_menu.css" rel="stylesheet" type="text/css" />
		
		<!--[if lt IE 7]><script type="text/javascript" src="includes/js/iehoverfix.js"></script><![endif]-->
	</head>
	<body>

		<!--#include virtual="menu/includes/ha_header.inc"-->

		<!--#include virtual="menu/includes/ha_menu.inc"-->

		<div id="content">
			<div id="columnA">
			<h2>Vendors</h2>
			<td width='25%' align='center'>
			<table width='100%' border='0' cellspacing='5' cellpadding='10'>

				<tr>
					<td align="center">
									<a href="https://www.uline.com">
									<img border="0" width='150' src="<%=strPathMenuImages%>uline.gif"></a><br>
						<a href='https://www.uline.com'>
						ULine</a>
					</td>
					<td align="center">
									<a href="http://xpedxstores.com/">
									<img border="0" src="<%=strPathMenuImages%>xpedx-logo.gif"></a><br>
						<a href='http://xpedxstores.com/'>
						xpedx</a><br>
					</td>
				</tr>
				<tr>
					<td colspan="2"align="center">
									<a href="http://www.prestigebox.com/index.html">
									<img border="0" src="<%=strPathMenuImages%>prestigebox.jpg"></a><br>
						<a href='http://www.prestigebox.com/index.html'>
						Prestige Box</a>
					</td>
				</tr>

			</table>
			<h2>Marketing</h2>
			<td width='25%' align='center'>
			<table width='100%' border='0' cellspacing='5' cellpadding='10'>

				<tr>
					<td align="center">
									<a href="https://adwords.google.com/select/Login?sourceid=awo&subid=ww-en-et-ads1_e&hl=en_us">
									<img border="0" src="<%=strPathMenuImages%>google_ad.gif"></a><br>
						<a href='https://adwords.google.com/select/Login?sourceid=awo&subid=ww-en-et-ads1_e&hl=en_us'>
						Google AdWords</a><br>
						<a href='https://www.google.com/webmasters/tools/siteoverview?hl=en&msg=2&pli=1'>
						Sitemaps</a>
					</td>
					<td align="center">
									<a href="https://login22.marketingsolutions.yahoo.com">
									<img border="0" src="<%=strPathMenuImages%>yahoo.gif"></a><br>
						<a href='https://login22.marketingsolutions.yahoo.com'>
						Yahoo</a>
					</td>
				</tr>
				<tr>
					<td colspan="2" align="center">
									<a href="http://ConstantContact.com">
									<img border="0" src="<%=strPathMenuImages%>ccontact.gif"></a><br>
						<a href='http://ConstantContact.com'>
						Constant Contact</a>
					</td>
				</tr>

			</table></div>
		<div id="columnB">

			<!--#include virtual="menu/includes/ha_shipping_holidays.inc"-->

			<!--#include virtual="menu/includes/ha_domains.inc"-->

		</div>
		<div style="clear: both;">&nbsp;</div>
		</div>

		<!--#include virtual="menu/includes/ha_footer.inc"-->

	</body>
</html>
