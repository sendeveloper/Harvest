<!DOCTYPE html>
<html>
<head>

  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/Config.asp"-->

<%
	ColorTab = 0
	PageHeading = "Home"
%>

    <title>Harvest American Backoffice - <%=PageHeading%></title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

    <link href="<%=strPathIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet">
    <link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>menu/dropdowntabfiles/halfmoontabs.css">

    <script type="text/javascript" src="<%=strPathIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>

    <style type="text/css">
    </style>

    <script language="javascript" type = "text/javascript">
    </script>
</head>

<body class="gray_desktop">
  <table class="MainBodyFixed">	
    <tr valign="top">
      <td>
        <!--#include virtual="ha_Backoffice/includes/bodyparts/ha_PageHeading.inc"-->
        <table class="MainTable">
			
<%
	For i = 1 to 15
		Response.Write "<tr><td>&nbsp;</td></tr>"
	Next
%>

		  <tr valign="top">
			<td align="center">
			  <table width='350' border='0' cellspacing='2' cellpadding='2' style="border: 1px solid black;">
				<tr>
				  <td align="left" style="font-size: 11px;">
					<span style="font-weight: bold; color: red;">
					  Warning:
					</span> 
					These facilities are solely for the use of authorized employees
					or agents of the organization and affiliates.  Unauthorized
					use is prohibited and subject to international criminal and civil penalties.  Individuals
					using this computer system are subject to having all of their activities on
					this system monitored and recorded by systems personnel.
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

