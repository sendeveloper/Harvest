<!--#include virtual="ha_backoffice/includes/Config.asp"-->
<!--#include virtual="includes/connection.asp"-->

<%
	ColorTab = 2
	PageHeading = "Telephone VPN Connections"
%>

<html>
<head>
    <title>Harvest American Backoffice - Telephone VPN Connections</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />

    <link href="<%=strPathBOIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet" />
    <link rel="stylesheet" type="text/css" href="<%=strPathBOIncludes%>menu/dropdowntabfiles/halfmoontabs.css" />

    <script type="text/javascript" src="<%=strPathBOIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>

    <style type="text/css">
		th
		{
		  text-align :       center;
          vertical-align :   top;
          background-color : #008000;
          font-weight :      bold;
          color :            white;
		}

		td.Heading1
		{
		  font-size:		14px;
		  text-align:		center;
		  font-weight:      bold;
		  border-bottom:	1px solid black;
		}
		lnk
		{
		  margin-left:      155;
		  text-align:		right;
          font-weight:		bold;
		}
    </style>

</head>

<body class="gray_desktop">
  <table align="center" cellspacing="0" cellpadding="0" class="MainBody">
    <tr>
	  <td>

		<!--#include virtual="ha_Backoffice/includes/bodyparts/ha_PageHeading.inc"-->

		<lnk><a href='../Company/DocumentMaintenance/ha_Document_download.asp?ID=1412'> VPN Diagram (Visio Doc)</a><lnk><br>
		
		<lnk><a href='../Company/DocumentMaintenance/ha_Document_download.asp?ID=1150'> How to Change a VPN IP Address (PDF)</a><lnk>
	
		<br><lnk><a href='/ha_Backoffice/HARVAMER.zip'> How to Change a VPN IP Address (Zipped Video)</a></lnk></br>
		
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
	
		  <tr><td>&nbsp;</td></tr>
		
		  <tr>
		    <td align="center">
			  <table width="75%" border="0" cellspacing="5" cellpadding="5">

				<tr>
				  <th>Location</th>
				  <th>Router LAN IP</th>
				  <th>Router Username</th>
				  <th>Router Password</th>
				  <th>WAN IP / Dynamic DNS</th>
				</tr>
				
				<tr>
				  <td>Camden (new)</td>
				  <td><a href="http://192.168.2.1" target="_vpn2b">192.168.2.1</a></td>
				  <td>davewj2o</td>
				  <td>get2itnow</td>
				  <td><a href="http://24.103.218.146" target="_wan2b">24.103.218.146</a></td>
				</tr>
				
				<!--<tr>
				  <td>Camden T1</td>
				  <td><a href="http://192.168.4.1" target="_vpn2a">192.168.4.1</a></td>
				  <td>davewj2o</td>
				  <td>get2itnow</td>
				  <td><a href="https://66.155.174.50/cgi-bin/welcome.cgi" target="_wan2a">66.155.174.50</a></td>
				</tr>-->
				
				<tr>
				  <td>Dave</td>
				  <td><a href="http://192.168.1.1" target="_vpn1">192.168.1.1</a></td>
				  <td>davewj2o</td>
				  <td>get2it</td>
				  <td><a href="http://davewj2o.dyndns.org" target="_wan1">davewj2o.dyndns.org</a></td>
				</tr>

				<tr>
				  <td>Teresa</td>
				  <td><a href="http://192.168.5.1" target="_vpn3">192.168.5.1</a></td>
				  <td>HarvestAmerican</td>
				  <td>get2it2day</td>
				  <td><a href="http://69.201.14.193" target="_wan3">69.201.14.193</a></td>
				</tr>

				<tr><td>&nbsp;</td></tr>

			  </table>
			</td>
		  </tr>
		  
		</table>
	  </td>
	</tr>
  </table>
  <!--#include virtual="ha_BackOffice/includes/BodyParts/ha_PageTail.inc"-->
</body>
</html>

