<%
	Session("Redirect") = ""
	ColorTab = 3
	PageHeading = "Servers Summary"
%>

<!--#include virtual="includes/connection.asp"-->
<!--#include virtual="ha_backoffice/includes/Config.asp"-->

<html>
<head>
    <title>Harvest American Backoffice - Servers</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

    <link href="<%=strPathIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet">
    <link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>menu/dropdowntabfiles/halfmoontabs.css">

    <script type="text/javascript" src="<%=strPathIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>

    <style type="text/css">

	td
		{
		font-family: Arial, Helvetica, sans-serif; 
		font-size: 12px; 
		font-style: normal; 
		line-height: normal; 
		font-weight: normal; 
		font-variant: normal; 
		text-transform: none;
		color: #000000;
		text-decoration: none;
		}

	table.ServerFarm
		{
		border: 2px solid black;
		background-color: #FFFFCC;
		width: 100%;
		height: 300px;
		}


	table.Server
		{
		border: 5px inset black;
		background-color: #DDDDDD;
		width: 100%;
		}

	td.ServerName
		{
		font-size: 14px; 
		font-weight: bold;
		}

	tr.ServerType
		{
		font-size: 12px; 
		text-align: center;
		}

	tr.ServerIP
		{
		font-size: 12px; 
		text-align: center;
		}

    </style>

    <script language="javascript" type = "text/javascript">

		function clickProperties(ServerID)
			{
			var URL = '<%=strPathControlPanel%>cp_ServerProperties.asp' +
				'?ServerID=' + ServerID;
			openPopUp(URL);
			}

		function clickPing(ServerID)
			{
			var URL = '<%=strPathControlPanel%>cp_ServerPing.asp' +
				'?ServerID=' + ServerID;
			openPopUp(URL);
			}

    </script>
</head>

<body class="gray_desktop">
  <table align="center" cellspacing="0" cellpadding="0" class="MainBody">
    <tr>
	  <td>

		<!--#include virtual="ha_Backoffice/includes/bodyparts/ha_PageHeading.inc"-->
		
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
	
		<tr>
		  <td align="center">
			<table width="90%" border="0" cellspacing="5" cellpadding="5">

			  <tr vAlign="top">
				<td width="250">
				  <a href="http://192.168.1.1"><h2>Home Office</h2></a>
				  <table cellspacing="5" cellpadding="5" class="ServerFarm">

					<tr>
					  <td>
						<table cellspacing="2" cellpadding="2" class="Server">
						  <tr>
							<td class="ServerName">
							  Corn
							</td>
						  </tr>
						  <tr class="ServerType">
							<td>
							  File Server
							</td>
						  </tr>
						  <tr class="ServerIP">
							<td>
							  McConnellsville
							</td>
						  </tr>
						</table>             
					  </td>
					</tr>

					<tr>
					  <td>
						<table cellspacing="2" cellpadding="2" class="Server">
						  <tr>
							<td class="ServerName">
							  Rye
							</td>
						  </tr>
						  <tr class="ServerType">
							<td>
							  SQL Database
							</td>
						  </tr>
						  <tr class="ServerIP">
							<td>
							  McConnellsville
							</td>
						  </tr>
						</table>             
					  </td>
					</tr>

					<tr>
					  <td height="100%">
						&nbsp;
					  </td>
					</tr>

				  </table>
				</td>

				<td width="300" style="visibility: hidden;">
				  <h2>Server Intellect - Dallas, TX</h2>
				  <table cellspacing="5" cellpadding="5" class="ServerFarm">

					<tr>
					  <td>
						<table cellspacing="2" cellpadding="2" class="Server">
						  <tr>
							<td class="ServerName">
							  Zippy01
							</td>
						  </tr>
						  <tr class="ServerType">
							<td>
							  Zip2Tax Production Server
							</td>
						  </tr>
						  <tr class="ServerIP">
							<td>
							  <%=LinkLine("Zippy01")%>
							</td>
						  </tr>
						</table>             
					  </td>
					</tr>

					<tr>
					  <td>
						<table cellspacing="2" cellpadding="2" class="Server">
						  <tr>
							<td class="ServerName">
							  Zippy02
							</td>
						  </tr>
						  <tr class="ServerType">
							<td>
							  Zip2Tax Production Server
							</td>
						  </tr>
						  <tr class="ServerIP">
							<td>
							  <%=LinkLine("Zippy02")%>
							</td>
						  </tr>
						</table>             
					  </td>
					</tr>

					<tr>
					  <td height="100%">
						&nbsp;
					  </td>
					</tr>

				  </table>
				</td>

				<td width="250">
				  <a href="http://192.168.2.1"><h2>Back-up Servers</h2></a>
				  <table cellspacing="5" cellpadding="5" class="ServerFarm">

					<tr>
					  <td>
						<table cellspacing="2" cellpadding="2" class="Server">
						  <tr>
							<td class="ServerName">
							  Camden1
							</td>
						  </tr>
						  <tr class="ServerType">
							<td>
							  SQL Database
							</td>
						  </tr>
						  <tr class="ServerIP">
							<td>
							  <%=LinkLine("Camden1")%>
							</td>
						  </tr>
						</table>             
					  </td>
					</tr>

					<tr>
					  <td>
						<table cellspacing="2" cellpadding="2" class="Server">
						  <tr>
							<td class="ServerName">
							  Camden2
							</td>
						  </tr>
						  <tr class="ServerType">
							<td>
							  FTP Server
							</td>
						  </tr>
						  <tr class="ServerIP">
							<td>
							  <%=LinkLine("Camden2")%>
							</td>
						  </tr>
						</table>             
					  </td>
					</tr>


					<tr>
					  <td>
						<table cellspacing="2" cellpadding="2" class="Server">
						  <tr>
							<td class="ServerName">
							  Camden3
							</td>
						  </tr>
						  <tr class="ServerType">
							<td>
							  Phone Log & File Backup Server
							</td>
						  </tr>
						  <tr class="ServerIP">
							<td>
							  <%=LinkLine("Camden3")%>
							</td>
						  </tr>
						</table>             
					  </td>
					</tr>

				  </table>
				</td>
			  </tr>

			</table>
		  </td>
		</tr>

		<tr>
		  <td>
			&nbsp;
		  </td>
		</tr>

		<tr>
		  <td align="center">
			<table width="90%" border="0" cellspacing="5" cellpadding="5">

			  <tr>
				<td width="250">
				  <h2>Active Servers - Spokane, WA</h2>
				  <table cellspacing="5" cellpadding="5" class="ServerFarm">

					<tr>
					  <td>
						<table cellspacing="2" cellpadding="2" class="Server">
						  <tr>
							<td class="ServerName">
							  Viking
							</td>
						  </tr>
						  <tr class="ServerType">
							<td>
							  Web Server
							</td>
						  </tr>
						  <tr class="ServerIP">
							<td>
							  <%=LinkLine("Viking")%>
							</td>
						  </tr>
						</table>             
					  </td>
					</tr>

					<tr>
					  <td height="100%">
						&nbsp;
					  </td>
					</tr>

				  </table>
				</td>

				<td width="250">
				  <a href="https://68.178.202.54:8443/login.php3"><h2>GoDaddy - Scottsdale, AZ</h2></a>
				  <table cellspacing="5" cellpadding="5" class="ServerFarm">

					<tr>
					  <td>
						<table cellspacing="2" cellpadding="2" class="Server">
						  <tr>
							<td class="ServerName">
							  Barley
							</td>
						  </tr>
						  <tr class="ServerType">
							<td>
							  Web Server - SQL Database
							</td>
						  </tr>
						  <tr class="ServerIP">
							<td>
							  <%=LinkLine("Barley")%>
							</td>
						  </tr>
						</table>
					  </td>
					</tr>

					<tr>
					  <td>
						<table cellspacing="2" cellpadding="2" class="Server">
						  <tr>
							<td class="ServerName">
							  Barley2
							</td>
						  </tr>
						  <tr class="ServerType">
							<td>
							  SQL Database - MySQL Database
							</td>
						  </tr>
						  <tr class="ServerIP">
							<td>
							  <%=LinkLine("Barley2")%>
							</td>
						  </tr>
						</table>
					  </td>
					</tr>

					<tr>
					  <td height="100%">
						&nbsp;
					  </td>
					</tr>

				  </table>
				</td>

				<td width="250">
				  <a href="https://orbit.theplanet.com/Login.aspx?url=/Network/OrbitGraphDetail.aspx?id=155931"><h2>SoftLayer - Dallas, TX (formerly The Planet)</h2></a>
				  <table cellspacing="5" cellpadding="5" class="ServerFarm">

					<tr>
					  <td>
						<table cellspacing="2" cellpadding="2" class="Server">
						  <tr>
							<td class="ServerName">
							  Casper10
							</td>
						  </tr>
						  <tr class="ServerType">
							<td>
							  SQL Database - Harvest Backoffice
							</td>
						  </tr>
						  <tr class="ServerIP">
							<td>
							  <%=LinkLine("Casper10")%>
							</td>
						  </tr>
						</table>             
					  </td>
					</tr>

					<tr style="visibility: hidden;">
					  <td>
						<table cellspacing="2" cellpadding="2" class="Server">
						  <tr>
							<td class="ServerName">
							  Houston01
							</td>
						  </tr>
						  <tr class="ServerType">
							<td>
							  MySQL Database
							</td>
						  </tr>
						  <tr class="ServerIP">
							<td>
							  <%=LinkLine("Houston01")%>
							</td>
						  </tr>
						</table>             
					  </td>
					</tr>

					<tr>
					  <td height="100%">
						&nbsp;
					  </td>
					</tr>

				  </table>
				</td>
			  </tr>
			</table>
		  </td>
		</tr>
		</table>
			
	  </td>
	</tr>
    <tr>
      <td>
        <div style="margin-left: 5em; height: 4em"><!--#include virtual="ha_BackOffice/includes/BodyParts/copyright.asp"--></div>
      </td>
    </tr>
  </table>
</body>
</html>

<%

Function LinkLine(servername)

    LinkLine = ""

    set rs=Server.CreateObject("ADODB.Recordset")

    'Read Properties
    sql = "SELECT Count(*) c from ha_Types WHERE Class = '" & servername & " - Properties'"
    rs.open sql, objcon, 1, 2
    if not rs.eof then
        if cInt(rs("c") > 0) then
            LinkLine = LinkLine & "<a href=""javascript:clickProperties('" & servername & " - Properties')"" class='Button50'>Properties</a>&nbsp;"
        end if
    end if
    rs.close

    'Read IP
    sql = "SELECT Value from ha_Types WHERE Class = '" & servername & " - Properties' AND Description = 'IP'"
    rs.open sql, objcon, 1, 2
    if not rs.eof then
        LinkLine = LinkLine & "<a href=""javascript:clickPing('" & rs("Value") & "')"" class='Button50'>Ping IP</a>&nbsp;"
    end if
    rs.close

    'Read Domain Name
    sql = "SELECT Value from ha_Types WHERE Class = '" & servername & " - Properties' AND Description = 'Domain Name'"
    rs.open sql, objcon, 1, 2
    if not rs.eof then
        LinkLine = LinkLine & "<a href=""javascript:clickPing('" & rs("Value") & "')"" class='Button50'>Ping Domain</a>&nbsp;"
    end if
    rs.close

    'Additional Services
    Select Case servername
      Case "Camden1" 
        LinkLine = LinkLine & "<a href=""" & strPathControlPanel & "cp_Services_last.asp"" class='Button50'>Services</a>&nbsp;"
      Case "Camden3"
        LinkLine = LinkLine & "<a href=""" & strPathTelephones & "ha_TelephoneCallLog.asp"" class='Button50'>Phone Log</a>&nbsp;"
    End Select

End Function

%>

