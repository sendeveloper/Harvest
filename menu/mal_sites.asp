<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<%
     Session("redirect") = ""

     ThisPage =" https://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL")
%>

<!--#include virtual="includes/connection.asp"-->

	<head>
		<meta http-equiv="content-type" content="text/html; charset=utf-8" />
		<title>Harvest American, Inc. | Sites & Servers</title>
		<link href="<%=strPathMenuIncludes%>ha_menu.css" rel="stylesheet" type="text/css" />
		<link href="<%=strPathMenuIncludes%>lightbox.css" rel="stylesheet" type="text/css" media="screen" />
		
		<!--[if lt IE 7]><script type="text/javascript" src="includes/js/iehoverfix.js"></script><![endif]-->
		<script type="text/javascript" src="<%=strPathMenuIncludes%>js/prototype.js"></script>
		<script type="text/javascript" src="<%=strPathMenuIncludes%>js/scriptaculous.js?load=effects,builder"></script>
		<script type="text/javascript" src="<%=strPathMenuIncludes%>js/lightbox.js"></script>
		<script type="text/javascript" src="<%=strPathMenuIncludes%>js/webcamRefresh.js"></script>
	</head>
	<body>

		<!--#include file="includes/ha_header.inc"-->

		<!--#include file="includes/ha_menu.inc"-->

		<div id="content">
		<div id="columnA">
			<h2>Sites & Servers</h2>
			<table width="80%" border="0" cellspacing="5" cellpadding="5"align="center">
				<tr valign="top">
					<table width="100%" border="0" cellspacing="5" cellpadding="9">
						<tr>
							<td align="center">
								<a href="http://www.HarvestAmerican.com">
								<img border="0" src="<%=strPathMenuImages%>logo_harvest.gif" height="65"></a><br>
								<a href="http://www.HarvestAmerican.com">HarvestAmerican.com</a><br>
								<font size="1">
								<a href="http://www.HarvestAmerican.com/stats/index.asp">Stats</a>
								</font>-
								<font size="1">
								<a href="http://www.HarvestAmerican.com/stats/servers.asp">Servers</a>
								</font>-
								<font size="1">
								<a href="<%=strPathBase%>/ControlPanel/cp_Servers.asp">CPanel</a><br>
								</font>
								<font size="1">
								<a href="http://www.harvestamerican.com/Wholesale/Tickets/">Wholesale</a>
								</font>									
							</td>
<!-- Use Eli"s Login -->
<%
   if trim(Session("ha_login")) = "erobbins" then
%>
							<td align="center">
								<a href="http://www.zip2tax.com/backoffice/z2t_Login.asp?lname=erobbins&pass=069780">
								<img border="0" alt="" src="<%=strPathMenuImages%>logo_z2t_sm.png" width="178" /></a><br>
								<table width="95%" border="0" cellspacing="1" cellpadding="1">
								  <tr>
									<td width="33%" align="center"><a href="http://www.zip2tax.com">Zip2Tax.com</a></td>
									<td width="33%" align="center"><a href="http://www.zip2tax.com/backoffice/z2t_Login.asp?lname=erobbins&pass=069780">Old Backoffice</a></td>
                                    <td width="33%" align="center"><a href="https://www.zip2tax.info/z2t_Backoffice/z2t_Login.asp?lname=erobbins&pass=069780">New Backoffice</a></td>
								 </tr>
								 <tr>
								 	<td width="50%" align="center" colspan="3"><font size="1"><a href="http://www.zip-codes.com">ZIP-Codes.com</a></font></td>
								 	<!--<td width="50%" align="center"><font size=1><a href="dl/Month End Procedures.txt">Web / Link</a></font></td>-->
								 </tr>
								</table>
							</td>
<!-- Use Jewel"s Login -->
<%
   elseif trim(Session("ha_login")) = "jlowe" then
%>
							<td align="center">
								<a href="http://www.zip2tax.com/backoffice/z2t_Login.asp?lname=jlowe&pass=090197">
								<img border="0" src="<%=strPathMenuImages%>logo_z2t_sm.png" width="178" /></a><br />
								<table width="95%" border="0" cellspacing="1" cellpadding="1">
								  <tr>
									<td width="33%" align="center"><a href="http://www.zip2tax.com">Zip2Tax.com</a></td>
									<td width="33%" align="center"><a href="http://www.zip2tax.com/backoffice/z2t_Login.asp?lname=jlowe&pass=090197">Old Backoffice</a></td>
                                    <td width="33%" align="center"><a href="https://www.zip2tax.info/z2t_Backoffice/z2t_Login.asp?lname=jlowe&pass=090197">New Backoffice</a></td>
								 </tr>
								 <tr>
								 	<td width="50%" align="center" colspan="3"><font size=1><a href="http://www.zip-codes.com">ZIP-Codes.com</a></font></td>
								 	<!--<td width="50%" align="center"><font size=1><a href="dl/Month End Procedures.txt">Web / Link</a></font></td>-->
								 </tr>
								</table>
							</td>
<!-- Use Lucinda"s Login -->
<%
   elseif trim(Session("ha_login")) = "lrowlands" then
%>
							<td align="center">
								<a href="http://www.zip2tax.com/backoffice/z2t_Login.asp?lname=lrowlands&pass=sequ0ia;">
								<img border="0" src="<%=strPathMenuImages%>logo_z2t_sm.png" width="178"></a><br>
								<table width="95%" border="0" cellspacing="1" cellpadding="1">
								  <tr>
									<td width="33%" align="center"><a href="http://www.zip2tax.com">Zip2Tax.com</a></td>
									<td width="33%" align="center"><a href="http://www.zip2tax.com/backoffice/z2t_Login.asp?lname=lrowlands&pass=sequ0ia;">Old Backoffice</a></td>
                                    <td width="33%" align="center"><a href="https://www.zip2tax.info/z2t_Backoffice/z2t_Login.asp?lname=lrowlands&pass=sequ0ia;">New Backoffice</a></td>
								 </tr>
								 <tr>
								 	<td width="50%" align="center" colspan="3"><font size=1><a href="http://www.zip-codes.com">ZIP-Codes.com</a></font></td>
								 	<!--<td width="50%" align="center"><font size=1><a href="dl/Month End Procedures.txt">Web / Link</a></font></td>-->
								 </tr>
								</table>
							</td>							
<!-- Anyone Else -->
<%
   else
%>
							<td align="center">
								<a href="http://www.zip2tax.com/backoffice/z2t_login.asp">
								<img border="0" src="<%=strPathMenuImages%>logo_z2t_sm.png" width="178"></a><br>
								<table width="95%" border="0" cellspacing="1" cellpadding="1">
								  <tr>
									<td width="33%" align="center"><a href="http://www.zip2tax.com">Zip2Tax.com</a></td>
									<td width="33%" align="center"><a href="http://www.zip2tax.com/backoffice/z2t_login.asp">Old Backoffice</a></td>
                                    <td width="33%" align="center"><a href="https://www.zip2tax.info/z2t_Backoffice/z2t_Login.asp">New Backoffice</a></td>
								 </tr>
								 <tr>
								 	<td width="50%" align="center" colspan="3"><font size=1><a href="http://www.zip-codes.com">ZIP-Codes.com</a></font></td>
								 	<!--<td width="50%" align="center"><font size=1><a href="dl/Month End Procedures.txt">Web / Link</a></font></td>-->
								 </tr>
								</table>
							</td>
<%
   end if
%>							
						</tr>
					</table>
				</tr>
				<tr valign="top">
					<table width="100%" border="0" cellspacing="5" cellpadding="9">
						<tr>
							<td align="center">
								<a href="http://www.NYSThruway.com">
								<img border="0" src="<%=strPathMenuImages%>nysthruway.png" width="75"></a><br>
								<a href="http://www.NYSThruway.com">NYSThruway.com</a>
							</td>
							<td align="center">
								<a href="http://www.EyeOnCamden.com">
								<img border="0" src="<%=strPathMenuImages%>eyeoncamden.jpg" height="65"></a><br>
								<a href="http://www.EyeOnCamden.com">EyeOnCamden.com</a><br /><font size=1><a href="http://www.EyeOnCamden.com/EyeonCamden/betasite/">New EyeOnCamden.com</a></font>
							</td>
						</tr>
					</table>
				</tr>
				<tr valign="top">
					<table width="100%" border="0" cellspacing="5" cellpadding="9">
						<tr>
							<td align="center" width="45%">
								<a href="https://68.178.202.54:8443/login.php3">
								<img border="0" src="<%=strPathMenuImages%>plesk.gif"></a><br>
								<a href="https://68.178.202.54:8443/login.php3">Plesk on Harvest</a>
							</td>
							<td align="center" width="55%">
								<a href="http://www.CamdenBudgetStorage.com/home/login/logpost.asp">
								<img border="0" src="<%=strPathMenuImages%>cbs.gif" height="25" width="178"></a><br>
								<a href="http://www.CamdenBudgetStorage.com/home/login/logpost.asp">Camden Budget Storage</a>
							</td>
						</tr>
					</table>
				</tr>
				<tr valign="top">
					<table style="width: 100%" border="0" cellspacing="5" cellpadding="9">
						<tr>
						  <td align="center">
							<a href="http://harvestamerican.hyperoffice.com">
							<img border="0" src="<%=strPathMenuImages%>hyperoffice_logo.gif" width="200" /></a><br />
								<a href="http://harvestamerican.hyperoffice.com">HyperOffice</a>
					      </td>
                          <td align="center">
                            <a href="https://www12.qth.com:2083/frontend/x3/index.html"><img border="0" src="http://qth.com/images/qthlogo.jpg" width="180" /></a><br />
                            <a href="https://www12.qth.com:2083/frontend/x3/index.html">Control Panel</a>&nbsp;&nbsp;<a href="https://www12.qth.com:2096/webmaillogout.cgi">Webmail</a>
                          </td>
						</tr>
					</table>
				</tr>
			</table>
		</div>
		
		<div id="columnB">

			<!--#include file="includes/ha_domains.inc"-->

            <!-- Webcam Selections -->
<%
   if trim(Session("ha_login")) = "erobbins" then
%>
            <!--#include virtual="menu/includes/ha_webcams_eli.inc"-->
<%
   else
%>
            <!--#include virtual="menu/includes/ha_webcams.inc"-->
<%
   end if
%>

		</div>
		<div style="clear: both;">&nbsp;</div>
		</div>

		<!--#include file="includes/ha_footer.inc"-->

	</body>
</html>
