<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

<!--#include virtual="includes/connection.asp"-->
<!--#include virtual="ha_backoffice/includes/Config.asp"-->
<!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->

<%
	Session("Redirect") = ""
	ColorTab = 4
	PageHeading = "Add/Remove Credentials"
	CheckLoadPermitted(10)
%>
 
<html>
	<head>
		<title>Harvest American Backoffice -  <%=PageHeading%></title>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

		<link href="<%=strPathIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet">
		<link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>menu/dropdowntabfiles/halfmoontabs.css">
		
		<script type="text/javascript" src="<%=strPathBOIncludes%>haFunctions.js"></script>
		<script language="javascript" type = "text/javascript">
		
			function clickPopup(ID)
				{
				var URL = '/ha_Backoffice/Company/DocumentMaintenance/ha_Document_Show.asp' +
					'?PhotoID=' + ID;
					void window.open(URL,'900','500')
				}
		 			
		</script>
		
		<script type="text/javascript" src="<%=strPathIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>
	 
		<style type="text/css">
	  
			span.subHead
				{
				font-weight: 	bold;
				color:			#336600;
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
				
				
			span.newStyleHeading
				{
				font-size: 		11pt;
				font-weight: 	bold;
				color:			#336600;
				}
				
	  </style>
		
	</head>

	<body class="gray_desktop">

	  <table align="center" cellspacing="0" cellpadding="0" class="MainBody">
	  
		<tr>
		
			<td>

				<!--#include virtual="ha_Backoffice/includes/bodyparts/ha_PageHeading.inc"-->
			
			</td>

		</tr>
		
		<tr>
		
			<td>
			
			<table width="1100px"  align="center">
				<tr> 
				
					<center>These are the procedures to be followed in order to add or remove an employee from the Harvest American system.</center><br>
				
				</tr>
				
				<tr>
				
					<td valign="top" width="47.5%" >
					
						<span class="newStyleHeading">Add</span><br>
						<hr>
						<span class="subHead">SQL</span><br>
						<a href="Javascript:clickPopup(1278)" class="titles">Info and Permissions</a><br>
						<a href="Javascript:clickPopup(1295)" class="titles">Accounts</a><br>
						<a href="Javascript:clickPopup(1279)" class="titles">Number-It Backoffice PopUps</a><br><br>
					
						<span class="subHead">E-mail</span><br>
						<a href="Javascript:clickPopup(1300)" class="titles">QTH Email</a><br>
						<a href="Javascript:clickPopup(1298)" class="titles">SmarterMail</a><br>
						<a href="Javascript:clickPopup(1299)" class="titles">Outlook</a><br>
						
					</td>
					
					<td width="5%">&nbsp;</td>
					
					<td valign="top" width="47.5%">
					
						<span class="newStyleHeading">Remove</span><br>
						<hr>
						<span class="subHead">SQL</span><br>
						<a href="Javascript:clickPopup(1337)" class="titles">Signature</a><br>
						<a href="Javascript:clickPopup(1338)" class="titles">Deactivate</a><br><br><br>
					
						<span class="subHead">E-mail</span><br>
						<a href="Javascript:clickPopup(1301)" class="titles">QTH</a><br>
						<a href="Javascript:clickPopup(1310)" class="titles">SmarterMail</a><br>
						
						
					</td>	
					
				</tr>
				
			</table>
				
			<%
			For i = 1 to 2
				Response.Write "<tr><td>&nbsp;</td></tr>"
			Next
			%>
			
		  </td>
		  
		</tr>
	
	</table>
  
	<!--#include virtual="ha_BackOffice/includes/BodyParts/ha_PageTail.inc"-->
  
	</body>

</html>

