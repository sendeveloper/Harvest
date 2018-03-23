  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsPagination.inc"-->
  
<%
    Response.buffer=true
    Response.clear

    Title = "Permission Definitions Information"

	Dim rs: Set rs = Server.CreateObject("ADODB.Recordset")
			
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

<html>
<head>
    <title><%=Title%></title>

    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <meta http-equiv="Cache-Control" content="no-cache">

    <link href="/ha_Backoffice/includes/ha_BackOfficePopup.css" type="text/css" rel="stylesheet">
    <script language="javascript" type="text/javascript" src="/includes/haFunctions.js"></script>

	<script type="text/javascript">
		function clickDownload(ID)
			{
			var URL = 'ha_Document_download.asp' +
				'?ID=' + ID;
			openPopUp(URL);
			}	
	</script>

	<style type="text/css">
		td
			{
			font-size: 10pt;
			}

		td.label
			{
			font-weight: bold;
			text-align: right;
			}
			
		td.head
			{
			font-weight: bold;
			color: #C0C0C0;
			border-bottom: 1px solid #C0C0C0;
			text-align: center;
			}
			
	</style
	
</head>


<body onLoad="SetScreen(900,700);">

  <span class="popupHeading"><%=Title%></span>

<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td>
	  
<%
		recordCount = 0
		LineCount = 0

		strSQL = "SELECT * FROM ha_PermissionDefinitions WHERE ID = (" & Request("id") & ")"
		rs.CursorLocation = 3
        rs.Open strSQL, connCasper10, 0, 1

%>

	  <table style="width: 90%; margin: 0 auto;" cellpadding="2" cellspacing="2">
		<tr><td>&nbsp;</td></tr>
		
		<tr valign="top">
		  <td class="label" colspan="3">
			ID:
		  </td>
		  <td colspan="3">
		    <%=rs("ID")%>
		  </td>
		</tr>
		
		<tr valign="top">
		  <td class="label" colspan="3">
			Program:
		  </td>
		  <td colspan="3">
		    <%=rs("Program")%>
		  </td>
		</tr>

		<tr valign="top">
		  <td class="label" colspan="3">
			Category:
		  </td>
		  <td colspan="3">
		    <%=rs("Category")%>
		  </td>
		</tr>

		<tr valign="top">
		  <td class="label" colspan="3">
			Page:
		  </td>
		  <td colspan="3">
		    <%=rs("Page")%>
		  </td>
		</tr>
		
		<tr valign="top">
		  <td class="label" colspan="3">
			Description:
		  </td>
		  <td colspan="3">
		    <%=rs("Description")%>
		  </td>
		</tr>


		<tr><td>&nbsp;</td></tr>
		

		
		<tr>
		  <td style="border-top: 1px solid #C0C0C0;" colspan="6">
		    &nbsp;
		  </td>
		</tr>
			
		
		<tr><td>&nbsp;</td></tr>
		
      </table>
		
    </td>
  </tr>
</table>

    <span class="popupButtons">
      <a href="javascript:window.close();" class="bo_Button80">Close</a>
    </span>

</body>
</html>
