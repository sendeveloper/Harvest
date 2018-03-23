<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<%
     Session("redirect") = ""

     ThisPage = "https://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL")
%>

<!--#include virtual="includes/connection.asp"-->

	<head>
		<meta http-equiv="content-type" content="text/html; charset=utf-8" />
		<title>Harvest American, Inc. | Misc</title>
		<link href="<%=strPathMenuIncludes%>ha_menu.css" type="text/css" rel="stylesheet" />
		<!--[if lt IE 7]><script type="text/javascript" src="includes/js/iehoverfix.js"></script><![endif]-->
	</head>
	<body>

		<!--#include virtual="menu/includes/ha_header.inc"-->

		<!--#include virtual="menu/includes/ha_menu.inc"-->

		<div id="content">
		<div id="columnA">
			<h2>Misc</h2>
			<table width="80%" border="0" cellspacing="5" cellpadding="5"align="center">
				<tr>
					<table width="100%" border="0" cellspacing="5" cellpadding="9">
						<tr>
							<td align="center">
								<a href="http://harvestamerican.hyperoffice.com/">
								<img border="0" src="<%=strPathMenuImages%>hyperoffice_logo.gif"></a><br>
								<a href="http://harvestamerican.hyperoffice.com/">HyperOffice</a>
							</td>
						</tr>
					</table>
				</tr>
				<tr>
					<table width="100%" border="0" cellspacing="5" cellpadding="9">
						<tr>
							<td align="center">
								<a href="http://www.youmail.com">
								<img border="0" src="<%=strPathMenuImages%>ymlogolove.gif"></a><br>
								<a href="http://www.youmail.com">YouMail:))</a>
							</td>
						</tr>
					</table>
				</tr>
				<tr>
					<table width="100%" border="0" cellspacing="5" cellpadding="9">
						<tr>
							<td align="center">
								<a href="http://www.w3schools.com">
								<img border="0" src="<%=strPathMenuImages%>w3schools.jpg"></a><br>
								<a href='http://www.w3schools.com'>W3 Schools</a>
							</td>
						</tr>
					</table>
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
