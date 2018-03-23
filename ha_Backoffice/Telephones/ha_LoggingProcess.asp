<!DOCTYPE html>
<!--#include virtual="ha_backoffice/includes/Config.asp"-->
<!--#include virtual="includes/connection.asp"-->

<%
	ColorTab = 2
	PageHeading = "Telephone Logging Process"
%>

<html>
<head>
	<title>Harvest American Backoffice - <%=PageHeading%></title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />

    <link href="<%=strPathBOIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet" />
    <link rel="stylesheet" type="text/css" href="<%=strPathBOIncludes%>menu/dropdowntabfiles/halfmoontabs.css" />

    <script type="text/javascript" src="<%=strPathBOIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>
	<script type="text/javascript" src="<%=strPathBOIncludes%>haFunctions.js" language="javascript"></script>
	
    <style type="text/css">
		span.subHead
			{
			font-weight: 	bold;
			color:			#336600;
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
					The Logging Process takes phone call data using a PERL script from the PBX System and transfers them to the Harvest Backoffice
					on Casper10.  Andrey Sergeev was the UpWork contractor who wrote this process.<br><br>
				  </td>
				</tr>

				<tr><td>&nbsp;</td></tr>
				  
				<tr>
				  <td align="right">
					<a href="javascript:clickPopupImage(1479);" class="TopRightLink">Telephone Logging Information</a>
					<a href="javascript:clickPopupImage(1481);" class="TopRightLink">Phone Log Upload Process</a>
				  </td>
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
