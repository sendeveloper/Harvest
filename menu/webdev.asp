<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<%
     Session("redirect") = ""

     ThisPage = "https://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL")
%>

<!--#include virtual="includes/connection.asp"-->

	<head>
		<meta http-equiv="content-type" content="text/html; charset=utf-8" />
		<title>Harvest American, Inc. | Web Dev</title>
		<meta name="keywords" content="" />
		<meta name="description" content="" />
		<link href="<%=strPathMenuIncludes%>ha_menu.css" rel="stylesheet" type="text/css" />
		
		<!--[if lt IE 7]><script type="text/javascript" src="includes/js/iehoverfix.js"></script><![endif]-->
	</head>
	<body>

		<!--#include virtual="menu/includes/ha_header.inc"-->

		<!--#include virtual="menu/includes/ha_menu.inc"-->

		<div id="content">
			<div id="columnA">
			<h2>Web Development</h2>
			<table width='100%' border='0' cellspacing='5' cellpadding='10'>

				<tr>
					<td align="center" width="55%">
						<a href="http://www.w3schools.com">
						<img border="0" src="<%=strPathMenuImages%>w3schools.jpg" height="75"></a><br />
						<a href="http://www.w3schools.com">W3 Schools</a>
					</td>
					<td align="center" width="45%">
						<a href="http://www.htmlhelp.com">
						<img border="0" src="<%=strPathMenuImages%>wdglogo.gif"></a><br />
						<a href="http://www.htmlhelp.com">Web Design Group</a>
					</td>
					</td>
				</tr>
				<tr>
					<td align="center">
						<a href="https://www.networksolutions.com/manage-it/index.jsp">
						<img border="0" src="<%=strPathMenuImages%>networksolutions.png" width="141"></a><br />
						<a href="https://www.networksolutions.com/manage-it/index.jsp">	Network Solutions</a>
					</td>
					<td align="center">
						<a href="https://www.siteground.com/customer_login.htm">
						<img border="0" src="<%=strPathMenuImages%>SiteHosting.gif"></a><br />
						<a href="https://www.siteground.com/customer_login.htm">SiteGround</a>
					</td>
				</tr>
			</table>
			<h2>Misc Links</h2>
			<table width='100%' border='0' cellspacing='5' cellpadding='10'>
						<tr>
							<td align="center">
								<a href="http://centralny.ynn.com/">
								<img border="0" src="<%=strPathMenuImages%>ynn_button.png"></a><br>
								<a href="http://centralny.ynn.com/">Your News Now</a>
							</td>
							<td align="center">
								<a href="http://www.youmail.com">
								<img border="0" src="<%=strPathMenuImages%>ymlogolove.gif"></a><br>
								<a href="http://www.youmail.com">YouMail:))</a>
							</td>							
						</tr>
			</table>
		</div>
		<div id="columnB">

			<!--#include virtual="menu/includes/ha_shipping_holidays.inc"-->

			<!--#include virtual="menu/includes/ha_domains.inc"-->

		</div>
		<div style="clear: both;">&nbsp;</div>
		</div>

		<!--#include virtual="menu/includes/ha_footer.inc"-->

	</body>
</html>
