<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->
  
<%
	Title = "Database Backup Purge Diagram"	
%>

<html>
<head>
  <title><%=Title%></title>

  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
  
  <link href="/ha_Backoffice/includes/ha_BackOfficePopup.css" type="text/css" rel="stylesheet">
  <script language="javascript" type="text/javascript" src="/includes/haFunctions.js"></script>
</head>

<body onLoad="SetScreen(900,900);">

	  <span class="popupHeading"><%=Title%></span>
	  
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">	  
  <tr>
    <td colspan="2">
	
      <table width="95%" border="0" cellspacing="2" cellpadding="2" align="center">

		<tr><td>&nbsp;</td></tr>
		
		<tr>
		  <td>
			For the first 2 weeks we keep the daily backups.  Then we fall back to once a week for
			2 months.  After 2 months we only keep once a month and archive it.
		  </td>
		</tr>		
		
		<tr>
		  <td align="center">
			<img src="/ha_BackOffice/Company/DocumentMaintenance/ha_Document_Show.asp?PhotoID=1112">
		  </td>
		</tr>
		
		<tr><td>&nbsp;</td></tr>
		
      </table>
    </td>
  </tr>
</table>

    <span class="popupCredit"><%=CreditLine(eCreatedBy, eCreatedDate, eEditedBy, eEditedDate)%></span>	

    <span class="popupButtons" style="padding-bottom: 20px;">
      <a href="javascript:window.close();" class="bo_Button80">Close</a>
    </span>

</body>
</html>
