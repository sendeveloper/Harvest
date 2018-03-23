<!DOCTYPE html>
<!--#include virtual="includes/connection.asp"-->
<!--#include virtual="ha_BackOffice/includes/Config.asp"-->

<%
	ColorTab = 2
	PageHeading = "Telephone PBX Information"
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
					A PBX is an old telephone installers term which stands for Private Branch Exchange (private telephone switchboard).  Ours is a
					Panasonic KX-TDO100 VOIP (Voice Over IP).<br><br>
					
					Changes to the telephone system can be made by logging into the PBX Control Panel.  It is mostly used for adding/removing extensions
					and adjusting the phone menus.
				  </td>
				</tr>

				<tr><td>&nbsp;</td></tr>
				  
				<tr>
				  <td align="right">
					<a href="javascript:clickPopupImage(1475);" class="TopRightLink">Connecting to the PBX Control Panel</a>
					<a href="javascript:clickPopupImage(1477);" class="TopRightLink">PBX Advanced Features</a>
				  </td>
				</tr>
				  
				<tr><td>&nbsp;</td></tr>
				
				<tr><td><hr></td></tr>
				  
				<tr>
				  <td align="Center">
					<img src="TelephonePBX.jpg">
				  </td>
				</tr>
					
				<tr><td><hr></td></tr>
						
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

