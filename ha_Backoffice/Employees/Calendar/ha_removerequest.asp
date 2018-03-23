<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsHelp.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->
  <!--#include virtual="ha_backoffice/includes/Dates/ConvDate.inc"-->

<html>
<head>
  <%

  %>
  
  <title>Harvest American Backoffice - <%=PageHeading%></title>
  <!--#include file="sql.asp"-->
  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
  
  <link href="/ha_BackOffice/includes/datepick/jquery-ui.css" rel="stylesheet" />
  <link href="/ha_BackOffice/includes/ha_Backoffice.css" type="text/css" rel="stylesheet" />
  
  <link href="/ha_Backoffice/includes/ha_BackOfficePopup.css" type="text/css" rel="stylesheet">
  <script language="javascript" type="text/javascript" src="/includes/haFunctions.js"></script>

  <script type="text/javascript">


  </script>
  <style type="text/css">
  body {
  overflow: hidden; 
  }
  </style>
</head>

<body onLoad="SetScreen(350,150);">

<center>
</br>
Your request has been removed.<br>

<a href="javascript:window.close();" class="bo_Button80">Close</a>

</center>
<%  
Dim strSQL, result, rs: Set rs = Server.CreateObject("ADODB.Recordset")
	
	strSQL = "UPDATE ha_EmpTimeRequest SET IsApproved=0 WHERE ID='" & Request("id") & "'"
	rs.Open strSQL, connCasper10, 1, 3 
	
	If Request("id") <> 0 then

	connCasper10.Execute strSQL
	
	connCasper10.close
	

		
	End If
	
	
%>  
</body>
</html>
