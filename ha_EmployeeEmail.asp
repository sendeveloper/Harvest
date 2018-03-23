<!DOCTYPE html>
<!--#include virtual="includes/connection.asp"-->
<!--#include virtual="ha_BackOffice/includes/Config.asp"-->

<%
	ColorTab = 4
	PageHeading = "Employee E-mail"
%>

<html>
<head>
	<title>Harvest American Backoffice - <%=PageHeading%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    
	<link href="<%=strPathIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet" />
	<link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>menu/dropdowntabfiles/halfmoontabs.css" />
	<script type="text/javascript" src="<%=strPathIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>
    <script type="text/javascript" src="<%=strPathBOIncludes%>haFunctions.js"></script>
  
    <style type="text/css">
		td
			{
			vertical-align:	top;
			}
			
		th
			{
			text-align:		left;
			border-bottom:	1px solid black;
			}
			
		div.members
			{
			border:			1px solid gray;
			height:			5em;
			}
    </style>

  <script language='JavaScript'>

	function clickPopupImage(ID)
		{
		var URL = '/ha_BackOffice/Company/DocumentMaintenance/ha_Document_Show.asp' +
			'?PhotoID=' + ID;
			PopupCenter(URL,'','900','500');
		}

  </script>
	
</head>

<body class="gray_desktop">
  <table cellpadding="0" cellspacing="0" class="MainBody" style="margin: 20px auto 0;width: 1200px !important;">
    <tr>
      <td>
        <!--#include virtual="ha_Backoffice/includes/bodyparts/ha_PageHeading.inc"-->
      </td>
    </tr>
	
	<tr><td>&nbsp;</td></tr>
	
    <tr>
      <td align="center">
        <table width="95%" cellpadding="2" cellspacing="2">
		  <tr>
		    <td colspan="4">
			  Our E-mail comes from our Casper09 Server.<br/>
			</tr>
		  </tr>
		  
		  <tr><td>&nbsp;</td></tr>
		  <tr><td>&nbsp;</td></tr>
		  
		  <tr>
			<td align="right" style="font-weight: bold; font-size: 11pt;" colspan="4">
			  <a href="javascript:clickPopupImage(1516);">E-mail Migration</a>&nbsp;&nbsp;
			  <a href="javascript:clickPopupImage(1547);">Mail Hierarchy for Forwarding</a>&nbsp;&nbsp;
			  <a href="javascript:clickPopupImage(1535);">E-mail Login Information</a>&nbsp;&nbsp;
			  <a href="javascript:clickPopupImage(1536);">User Creation Tutorial</a>&nbsp;&nbsp;
			  <a href="javascript:clickPopupImage(1515);">Setting up Outlook</a>
			</td>
		  </tr>
		  

		  <tr><td>&nbsp;</td></tr>
		  		  
		</table>
      </td>
    </tr>
	
	<tr><td>&nbsp;</td></tr>
	<tr><td>&nbsp;</td></tr>

  </table>
  <!--#include virtual="ha_BackOffice/includes/BodyParts/ha_PageTail.inc"-->
</body>
<script type="text/javascript">
	alert("This is an alert box!");
	</script>
</html>

