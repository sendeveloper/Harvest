<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

<!--#include virtual="includes/connection.asp"-->
<!--#include virtual="ha_backoffice/includes/Config.asp"-->

<%
	Session("Redirect") = ""
	ColorTab = 4
	PageHeading = "Employee Setup Procedures"
	response.write "Before CheckLoadPermitted<br>"
	CheckLoadPermitted(10)
%>

<html>
<head>
    <title>Harvest American Backoffice -  <%=PageHeading%></title>
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
  <table align="center" cellspacing="0" cellpadding="0" class="MainBody">
    <tr>
	  <td>

	<!--#include virtual="ha_Backoffice/includes/bodyparts/ha_PageHeading.inc"-->
	
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

