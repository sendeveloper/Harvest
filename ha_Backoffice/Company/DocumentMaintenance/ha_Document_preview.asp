<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->
  <!--#include virtual="ha_backoffice/includes/Dates/ConvDate.inc"-->
  
<%
	Title = "Document Preview"	

	If isNull(Request("ID")) or Request("ID") = "" Then
		PhotoID = 0
		Message = "No Image ID"
	Else
		PhotoID = Request("ID")
	End If
	
	Set rs = Server.CreateObject("ADODB.Recordset")
    strSQL = "ha_Document_info(" & PhotoID & ")"
    rs.Open strSQL, connCasper10, 3, 3, 4	
	
    If rs.eof Then
		PhotoID = 0
		Message = "Record not found for this Image ID"
    Else
		'If rs("fType") <> "image/jpeg" Then
		'	'response.write rs("fType")
		'	PhotoID = 0
		'	Message = "Image not a jpg"
		'Else
			Message 		= Null2Space(rs("fName"))
			eCreatedBy 		= Null2Space(rs("OriginalCreatedBy"))
			eCreatedDate 	= Null2Space(rs("OriginalCreatedDate"))
			eEditedBy 		= Null2Space(rs("CreatedBy"))
			eEditedDate		= Null2Space(rs("CreatedDate"))
		'End If
    End If
%>

<html>
<head>
  <title><%=Title%></title>

  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
  
  <link href="/ha_Backoffice/includes/ha_BackOfficePopup.css" type="text/css" rel="stylesheet">
  <script language="javascript" type="text/javascript" src="/includes/haFunctions.js"></script>
</head>

<body onLoad="SetScreen(700,750);">

	  <span class="popupHeading"><%=Title%></span>
	  
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">	  
  <tr>
    <td colspan="2">
	
      <table width="95%" border="0" cellspacing="2" cellpadding="2" align="center">

		<tr><td>&nbsp;</td></tr>
		
		<tr>
		  <td align="center">
<%
	If PhotoID > 0 Then
		If rs("fType") = "app/pdf" Then
			%>
			<object src="ha_Document_Show.asp?PhotoID=<%=cstr(PhotoID)%>” type="application/pdf"></object>
			<!--<object width="500" height="300" type="appplication/pdf" src="ha_Document_Show.asp?PhotoID=<%=cstr(PhotoID)%>”></object>-->
			<%
		Else
			Response.write "<img src='ha_Document_Show.asp?PhotoID=" & cstr(PhotoID) & "' width='400'>"
		End If
	End If
%>
		  </td>
		</tr>
		
		<tr>
		  <td align="center">
		    <%=Message%>
		  </td>
		</tr>
		
		<tr><td>&nbsp;</td></tr>
		
      </table>
    </td>
  </tr>
</table>

    <span class="popupCredit"><%=CreditLine(eCreatedBy, eCreatedDate, eEditedBy, eEditedDate)%></span>	

    <span class="popupButtons" style="padding-bottom: 50px;">
      <a href="javascript:window.close();" class="bo_Button80">Close</a>
    </span>

</body>
</html>
