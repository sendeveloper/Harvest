<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<%
     Session("redirect") = ""

     ThisPage = "https://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL")
%>

<!--#include virtual="includes/connection.asp"-->

	<head>
		<meta http-equiv="content-type" content="text/html; charset=utf-8" />
		<title>Harvest American, Inc. | Accounting</title>
		<link href="<%=strPathMenuIncludes%>ha_menu.css" type="text/css" rel="stylesheet" />
		<!--[if lt IE 7]><script type="text/javascript" src="includes/js/iehoverfix.js"></script><![endif]-->
	</head>
	<body>

		<!--#include virtual="menu/includes/ha_header.inc"-->
		<!--#include virtual="menu/includes/ha_menu.inc"-->
		
		<div id="content">
		<div id="columnA">
			<h2>Accounting</h2>
			<table width="80%" border="0" cellspacing="5" cellpadding="5"align="center">
				<tr valign="top">
					<table width="100%" border="0" cellspacing="5" cellpadding="9">
						<tr>
							<td align="center">
								<a href="https://www.accessfcu.org/onlineserv/HB/Signon.cgi">
								<img border="0" src="<%=strPathMenuImages%>accessfcu.gif"></a><br>
								<a href='https://www.accessfcu.org/onlineserv/HB/Signon.cgi'>Access Federal CU</a>
							</td>
							<td align="center">
								<a href="https://www.paypal.com">
								<img border="0" src="<%=strPathMenuImages%>paypal.gif" height="63"></a><br>
								<a href='http://www.paypal.com'>PayPal</a>
							</td>
						</tr>
					</table>
				</tr>
				<tr valign="top">
					<table width="100%" border="0" cellspacing="5" cellpadding="9">
						<tr>
							<td colspan="2" align="center">
								<a href="http://www.bankofamerica.com">
								<img border="0" src="<%=strPathMenuImages%>boa.gif" height="100"></a><br>
								<a href='http://www.bankofamerican.com'>Bank of America</a>
							</td>

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
