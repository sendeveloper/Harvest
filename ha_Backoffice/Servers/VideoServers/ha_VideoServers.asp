<!DOCTYPE html>
<!--#include virtual="includes/connection.asp"-->
<!--#include virtual="ha_BackOffice/includes/Config.asp"-->

<%
	ColorTab = 1
	PageHeading = "Video Servers"
%>

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<title>Harvest American Backoffice - <%=PageHeading%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    
	<link href="<%=strPathIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet" />
	<link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>menu/dropdowntabfiles/halfmoontabs.css" />
	
	<script type="text/javascript" src="<%=strPathIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>
	<script type="text/javascript" src="<%=strPathBOIncludes%>haFunctions.js" language="javascript"></script>
  
    <style type="text/css">
		td
			{
			vertical-align:	top;
			}
			
		th
			{
			border-bottom:	1px solid black;
			}
			
		div.members
			{
			border:			1px solid gray;
			height:			7em;
			}
    </style>
</head>

<body class="gray_desktop">
  <table class="MainBodyFixed">	
    <tr valign="top">
      <td>
        <!--#include virtual="ha_Backoffice/includes/bodyparts/ha_PageHeading.inc"-->
        <table class="MainTable">
	
		  <tr><td>&nbsp;</td></tr>
	
		  <tr>
			<td align="center">
			  <table width="95%" cellpadding="2" cellspacing="2">
				<tr>
				  <td>
					Video Servers support our WebCams. All Camden video servers are now on LAN 2.
				  </td>
				</tr>

				<tr><td>&nbsp;</td></tr>
				  
				<tr>
				  <td align="right">
					<a href="javascript:clickPopupImage(1471);" class="TopRightLink">Standard Camera Setup</a>
				  </td>
				</tr>
								  
				<tr><td>&nbsp;</td></tr>
				  
				<tr>
				  <td align="Center">
					<table width="80%" cellpadding="2" cellspacing="2">
		  
					  <tr>
						<th width="30%">Location</th>
						<th width="30%">IP</th>
						<th width="10%">&nbsp;</th>
						<th width="10%">&nbsp;</th>				
						<th width="20%">&nbsp;</th>
					  </tr>
													
					  <tr>
						<td>Camden - Dave's office</td>
						<td>192.168.2.214</td>
					  </tr>
					
					  <tr>
						<td>Camden - Angel's Cube</td>
						<td>192.168.2.245</td>
					  </tr>
					  
					  <tr>
						<td>Camden - Server Room</td>
						<td>192.168.2.212</td>
					  </tr>
					
					  <tr>
						<td>McConnellsville</td>
						<td>192.168.1.101</td>
					  </tr>
					
					  <tr>
						<td colspan="5">
						  <hr>
						</td>
					  </tr>
				
					</table>
				  </td>
				</tr>
		
				<tr><td>&nbsp;</td></tr>
		  
				<tr>
				  <td colspan="4">
					<hr>
				  </td>
				</tr>
		  
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

