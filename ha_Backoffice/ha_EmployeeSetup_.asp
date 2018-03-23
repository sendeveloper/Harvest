<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

<!--#include virtual="includes/connection.asp"-->
<!--#include virtual="ha_backoffice/includes/Config.asp"-->

<%
	Session("Redirect") = ""
	ColorTab = 4
	PageHeading = "Employee Setup Procedures"
	CheckLoadPermitted(10)
%>

<html>
<head>
    <title>Harvest American Backoffice -  <%=PageHeading%></title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

    <link href="<%=strPathIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet">
    <link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>menu/dropdowntabfiles/halfmoontabs.css">
    <script language="javascript" type = "text/javascript">
	

function clickPopup(ID)
		{
		var URL = '../ha_Backoffice/Company/DocumentMaintenance/ha_Document_Show.asp' +
			'?PhotoID=' + ID;
			void window.open(URL,'900','500')
		}
	

		
    </script>
    <script type="text/javascript" src="<%=strPathIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>

  


	
	  <style type="text/css">
  
		a
			{
			font-size:		10pt;
			}
		
		span.subHead
			{
			font-size:		14pt;
			font-weight:	bold;
			
			}
		span.superHead
			{
			font-size:		18pt;
			font-weight:	bold;
			
			}
		.titles
			{
			padding-left:3em;
			}
			
		.innerTable
			{
			background-color: #ECF7F1;
			}
  </style>
	
</head>

<body class="gray_desktop">
  <table align="center" cellspacing="" cellpadding="" class="MainBody">
    <tr>
		<td>

		<!--#include virtual="ha_Backoffice/includes/bodyparts/ha_PageHeading.inc"-->
		</td>
	</tr>
	<tr>
		<td>
		<table class="innerTable" width="1200px">
			<tr>
				<td>
				<span class="superHead">Adding an Employee</span>
					<br></br><span class="subHead">SQL</span><br>
						<a href="Javascript:clickPopup(1286)" class="titles">Basic Employee Information and Permissions</a></br>
						<a href="Javascript:clickPopup(1286)" class="titles">Basic Employee Accounts</a></br>
						<a href="Javascript:clickPopup(1286)" class="titles">Number-It.info BackOffice PopUp Functionality</a>
				
					<br></br><span class="subHead">Email</span></br>
						<a href="Javascript:clickPopup(1286)" class="titles">Adding an Employee to QTH Email</a></br>
						<a href="Javascript:clickPopup(1286)" class="titles">Adding an Employee to SmarterMail</a></br>
						<a href="Javascript:clickPopup(1286)" class="titles">Adding an Employee to PLESK</a></br>
						<a href="Javascript:clickPopup(1286)" class="titles">Properly Setting Up an Account in Outlook</a></br>
				</td>
			<!--<td>
					<span class="superHead">Removing an Employee</span>
						<br></br><span class="subHead">SQL</span><br>
				</td>-->
			</tr>
			</table>
	
		
	
		<%
	For i = 1 to 35
		Response.Write "<tr><td>&nbsp;</td></tr>"
	Next
%>
	  </td>
	</tr>
  </table>
  <!--#include virtual="ha_BackOffice/includes/BodyParts/ha_PageTail.inc"-->
</body>
</html>

