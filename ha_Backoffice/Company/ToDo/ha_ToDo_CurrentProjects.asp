<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsPermissions.inc"-->
  
<%
	Title = "To Do Current Projects"
	
	Dim rs: Set rs = Server.CreateObject("ADODB.Recordset")

	strSQL = "ha_Employee_read_by_id(" & Request("ID") & ")"
	rs.Open strSQL, connLocal, 3, 3, 4
	If not rs.eof Then
		EmployeeNameFirst = Ltrim(Null2Space(rs("FirstName")))
		EmployeeNameShort = Ltrim(Null2Space(rs("ShortName")))
		EmployeeName = Ltrim(Null2Space(rs("FirstName")) & " " & Null2Space(rs("LastName")))
	End If
	rs.Close
%>

<html>
<head>
  <title><%=Title%></title>

  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
  
  <link href="/ha_Backoffice/includes/ha_BackOfficePopup.css" type="text/css" rel="stylesheet">
  <script language="javascript" type="text/javascript" src="/includes/haFunctions.js"></script>
  
  <style type="text/css">
	th
		{
		font-size:		11pt;
		border-bottom: 	1px solid black;
		}
  </style>

  <script type="text/javascript">
  
	var formLoaded = false;
	
	function changeAssignee()  
		{
		if (formLoaded == true)
			{
			var ID = document.getElementById("Assignee").value;
			var URL = 'ha_ToDo_CurrentProjects.asp' + 
				'?ID=' + ID;
			location.href = URL;
			}
		}
		
	function formLoad()
		{
		SetScreen(1000,850);
		formLoaded = true;
		}
		
  </script>
</head>

<body onLoad="formLoad();">

	  <span class="popupHeading"><%=Title%></span>
	  
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">	  
  <tr>
    <td>
	
      <table width="95%" border="0" cellspacing="2" cellpadding="1" align="center">
		<tr>
		  <td width="50%">
			Assignee
			<Select id="Assignee" name="Assignee" onchange="changeAssignee()">

<%
	strSQL = "util_HTML_option_list('ha_EmployeesFormatted_active', 'ProjectSpecialist', '1', 'ID', 'ShortName', 'FirstNameOrder', '', '" & Request("ID") & "')"
	rs.Open strSQL, connLocal, 3, 3, 4
	If not rs.eof Then
		While not rs.eof
			Response.write rs("result")
			rs.MoveNext
		Wend
	End If
	rs.Close
%>
		
			</Select>

			</td>
		  <td width="50%" align="right">
			<%=now()%>
		  </td>
		</tr>
		
		<tr><td>&nbsp;</td></tr>
		
		<tr>
		  <td colspan="2">
		    This is a list of To Do Projects that have a status of "In Progress" for <%=EmployeeName%>
		  </td>
		</tr>
		
		<tr><td>&nbsp;</td></tr>

	  </table>
		
      <table width="95%" border="0" cellspacing="2" cellpadding="1" align="center">
	  	  
  		<tr>
		  <th width="10%">
		    Priority
		  </th>
		  <th width="90%">
		    Title
		  </th>
		</tr>
		
		<tr><td>&nbsp;</td></tr>

	    <tr>
		  <td style="font-size: 10pt; font-style: italic;">
		    Assigned To:
		  </td>
		</tr>		
		
<%
					strSQL = "ha_ToDo_CurrentProjects('" & EmployeeNameShort & "', 'Assigned')"
					rs.Open strSQL, connLocal, 3, 3, 4
					If not rs.eof Then
						While not rs.eof
							'Skip a line if the priority value changes
							If rs("Priority") <> PreviousPriority Then
								PreviousPriority = rs("Priority")
								Response.write "<tr><td>&nbsp;</td></tr>"
							End If
%>
							
		<tr>
		  <td align="center">
		    <%=rs("Priority")%>
		  </td>
		  <td>
		    <%=rs("Title")%>
		  </td>
		</tr>
<%
							rs.MoveNext
						Wend
					End If
					rs.Close
%>
		<tr><td>&nbsp;</td></tr>

	    <tr>
		  <td style="font-size: 10pt; font-style: italic;">
		    Managed By:
		  </td>
		</tr>		
		
<%
					strSQL = "ha_ToDo_CurrentProjects('" & EmployeeNameShort & "', 'Managed')"
					rs.Open strSQL, connLocal, 3, 3, 4
					If not rs.eof Then
						While not rs.eof
							'Skip a line if the priority value changes
							If rs("Priority") <> PreviousPriority Then
								PreviousPriority = rs("Priority")
								Response.write "<tr><td>&nbsp;</td></tr>"
							End If
%>
							
		<tr>
		  <td align="center">
		    <%=rs("Priority")%>
		  </td>
		  <td>
		    <%=rs("Title")%>
		  </td>
		</tr>
<%
							rs.MoveNext
						Wend
					End If
					rs.Close
%>

	  </table>
					
      <table width="95%" border="0" cellspacing="2" cellpadding="1" align="center">
					
		<tr>
		  <td style="border-bottom: 2px solid gray;">
			&nbsp;
		  </td>
		</tr>
		
		<tr><td>&nbsp;</td></tr>
		
	  </table>
    </td>
  </tr>
</table>

    <span class="popupButtons" style="padding-bottom: 3em;">
      <a href="javascript:window.close();" class="bo_Button80">Close</a>
    </span>

</body>
</html>
