  <!--#include virtual="includes/connection.asp"-->
  
<%
    Response.buffer=true
    Response.clear

    Title = "Active Staff"

	Dim rs: Set rs = Server.CreateObject("ADODB.Recordset")
		
    SQL="ha_Employee_Active_list"
    rs.open SQL, connLocal, 3, 3, 4	
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

<html>
<head>
    <title><%=Title%></title>

    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

    <link href="/ha_Backoffice/includes/ha_BackOfficePopup.css" type="text/css" rel="stylesheet">
    <script language="javascript" type="text/javascript" src="/includes/haFunctions.js"></script>

	<style type="text/css">
		td
			{
			font-size: 10pt;
			}

		td.head
			{
			font-weight: bold;
			font-size: 8pt;
			font-stretch: condensed;
			border-bottom: 1px solid black;
			}
		td.total
			{
			font-weight: bold;
			font-size: 9pt;
			border-top: 1px solid black;
			}
		h2
			{
			font-size: 12pt;
			font-weight: bold;
			border-bottom: 2px solid #C0C0C0;
			}
		span.Manager1
			{
			font-size: 8pt;
			color: #FFFFFF;
			}
		span.Manager2
			{
			font-size: 8pt;
			color: #C0C0C0;
			}
	</style
	
</head>


<body onLoad="SetScreen(1100,850);">

  <span class="popupHeading"><%=Title%></span>

<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td>
	  
		<table width="90%" border="0" cellspacing="2" cellpadding="2" align="center">

			<!----------------- Full Time ----------------------->
			<tr><td>&nbsp;</td></tr>
			<tr><td colspan="12"><h2>Full Time</h2></td></tr>

			<tr>
			  <td class="head" width="15%">
				Name
			  </td>
			  <td class="head" width="15%">
				Location
			  </td>
			  <td class="head" width="70%">
				Job Description
			  </td>				
			</tr>

<%			
	rs.MoveFirst
	If Not rs.eof Then
		While Not rs.eof
			If rs("EmpStatus") = "Full Time" Then
%>
			<tr>
			  <td>
				<%=rs("Name")%>
			  </td>
			  <td>
				<%=rs("Location")%>
			  </td>
			  <td>
				<%=rs("JobTitle")%>
			  </td>
			</tr>
<%
			End If
			rs.MoveNext
		Wend
	End If
%>
			
			<tr><td>&nbsp;</td></tr>
			
			<!----------------- Part Time ----------------------->
			<tr><td>&nbsp;</td></tr>
			<tr><td colspan="12"><h2>Part Time</h2></td></tr>

			<tr>
			  <td class="head" width="12%">
				Name
			  </td>
			  <td class="head" width="12%">
				Location
			  </td>
			  <td class="head" width="76%">
				Job Description
			  </td>				
			</tr>
			
<%			
	rs.MoveFirst
	If Not rs.eof Then
		While Not rs.eof
			If rs("EmpStatus") = "Part Time" Then
%>
			<tr>
			  <td>
				<%=rs("Name")%>
			  </td>
			  <td>
				<%=rs("Location")%>
			  </td>
			  <td>
				<%=rs("JobTitle")%>
			  </td>
			</tr>
<%
			End If
			rs.MoveNext
		Wend
	End If
%>
			
			<tr><td>&nbsp;</td></tr>
			
			<!----------------- Contractors ----------------------->
			<tr><td>&nbsp;</td></tr>
			<tr><td colspan="12"><h2>Contractors</h2></td></tr>

			<tr>
			  <td class="head" width="12%">
				Name
			  </td>
			  <td class="head" width="12%">
				Location
			  </td>
			  <td class="head" width="76%">
				Job Description
			  </td>				
			</tr>
			
<%			
	rs.MoveFirst
	If Not rs.eof Then
		While Not rs.eof
			If rs("EmpStatus") = "Contractor" Then
%>
			<tr>
			  <td>
				<%=rs("Name")%>
			  </td>
			  <td>
				<%=rs("Location")%>
			  </td>
			  <td>
				<%=rs("JobTitle")%>
			  </td>
			</tr>
<%
			End If
			rs.MoveNext
		Wend
	End If
%>
			
			<tr><td>&nbsp;</td></tr>
			
		</table>
				
    </td>
  </tr>
  
  <tr>
    <td>
      <span class="popupButtons">
		<a href="javascript:window.close();" class="bo_Button80">Close</a>
      </span>
	</td>
  </tr>
	
  <tr><td>&nbsp;</td></tr>
	
</table>


</body>
</html>
