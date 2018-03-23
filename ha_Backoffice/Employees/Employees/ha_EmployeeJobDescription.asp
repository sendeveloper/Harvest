<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsDates.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsHelp.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsPermissions.inc"-->
  <!--#include virtual="ha_backoffice/includes/Dates/ConvDate.inc"-->
  
<%
	Dim rs: Set rs = Server.CreateObject("ADODB.Recordset")

	Title = "Employee Job Description"		
%>

<html>
<head>
  <title><%=Title%></title>

  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
  
  <link href="/ha_Backoffice/includes/ha_BackOfficePopup.css" type="text/css" rel="stylesheet">
  <script language="javascript" type="text/javascript" src="/includes/haFunctions.js"></script>
  
  <style type="text/css">
	.nameEmp
		{
		font-weight:	bold;
		font-size:		11pt;
		text-align:		left;
		width:			50%;
		}
		
	.nameDate
		{
		font-weight:	bold;
		text-align:		right;
		width:			50%;
		}
		
	.nameTitle
		{
		font-weight:	bold;
		font-size:		14pt;
		text-align:		center;
		color:			gray;
		border-bottom:	2px solid gray;
		}

	.namePurpose
		{
		font-weight:	bold;
		font-size:		10pt;
		color:			gray;
		}
		
	.nameTask
		{
		font-weight:	bold;
		}
	
	.GrayLine
		{
		border-bottom:	2px solid gray;
		}
		
	.GrayLineTop
		{
		border-top:		1px solid gray;
		}
</style>

</head>

<body onLoad="SetScreen(1000,850);">

	  <span class="popupHeading"><%=Title%></span>
	  
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">	  
  <tr>
    <td colspan="2">
	
      <table width="95%" border="0" cellspacing="2" cellpadding="1" align="center">

<%
					strSQL = "ha_Employee_JobDescription(" & Request("ID") & ")"
					rs.Open strSQL, connLocal, 3, 3, 4
					If not rs.eof Then
						While not rs.eof
							Response.write  rs("Result") & vbCrLf
							rs.MoveNext
						Wend
					End If
					rs.Close
%>
					
		<tr><td class="GrayLine">&nbsp;</td></tr>
		<tr>
		  <td style="color: gray; font-size: 11pt; font-weight: bold;">
		    Harvest American, Inc.
		  </td>
		</tr>
	  </table>
    </td>
  </tr>
</table>

    <span class="popupButtons">
      <a href="javascript:window.close();" class="bo_Button80">Close</a>
    </span>

</body>
</html>
