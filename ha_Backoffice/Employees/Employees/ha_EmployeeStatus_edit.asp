<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsDates.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsHelp.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->
  <!--#include virtual="ha_backoffice/includes/Dates/ConvDate.inc"-->
  
<%
	Dim rs: Set rs = Server.CreateObject("ADODB.Recordset")

		ID = Request("ID")
		Title = "Employee Status"
		
		strSQL = "SELECT * FROM ha_EmployeesFormatted WHERE ID = " & ID
        rs.open strSQL, connCasper10, 0, 1
		

		eFullName			= Null2Space(rs("FullName"))
		eHiredDate			= Null2Space(rs("HiredDate"))
		eTerminationDate	= Null2Space(rs("TerminationDate"))
		eStatus				= Null2Space(rs("EmpStatus"))
		eType				= Null2Space(rs("EmpType"))
		eContractor			= IIf(rs("Contractor")=True,"Yes", "No")
		eProjectSpecialist	= IIf(rs("ProjectSpecialist")=True,"Yes", "No")
		eTimeClock			= IIf(rs("PermitTimeClock")=True,"Yes", "No")
		eStaffStatus		= IIf(rs("PermitStaffStatus")=True,"Yes", "No")

		eCreatedBy			= Null2Space(rs("CreatedBy"))
		eCreatedDate		= Null2Space(rs("CreatedDate"))
		eEditedBy			= Null2Space(rs("EditedBy"))
		eEditedDate			= Null2Space(rs("EditedDate"))
		
        rs.close
		
%>

<html>
<head>
  <title><%=Title%></title>

  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
  
  <link href="/ha_Backoffice/includes/ha_BackOfficePopup.css" type="text/css" rel="stylesheet">
  <script language="javascript" type="text/javascript" src="/includes/haFunctions.js"></script>
</head>

<body onLoad="SetScreen(650,650);">
  <form method="Post" action="ha_EmployeeStatus_post.asp?ID=<%=ID%>" name="frm">

	  <span class="popupHeading"><%=Title%></span>
	  
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">	  
  <tr>
    <td colspan="2">
	
      <table width="95%" border="0" cellspacing="2" cellpadding="1" align="center">

		<tr><td width="35%>&nbsp;</td><td width="65%>&nbsp;</td></tr>
			
		<tr>
		  <td align="right" style="padding-bottom: 8px;">
			Name:
		  </td>
		  <td style="padding-bottom: 8px;">
		    <b><%=eFullName%></b>
		  </td>		 
		</tr>
		
		<tr>
		  <td align="right">
			Hire Date:
		  </td>
		  <td>
			<input type="text" name="HireDate" id="HireDate"
				size="10" value="<%=eHiredDate%>">
			<%=PopupDate("HireDate")%>
		  </td>		 
		</tr>
		
		<tr>
			<td align="right">
				Status:
			</td>
			<td>
				<Select id="EmpStatus" name="EmpStatus" style="width: 125px;">
  
<%
					strSQL = " util_HTML_option_list_from_table('ha_EmployeesFormatted', 'EmpStatus', '" & eStatus & "')"
					rs.Open strSQL, connCasper10, 3, 3, 4
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
		</tr>
		
		<tr>
			<td align="right">
				Type:
			</td>
			<td>
				<Select id="EmpType" name="EmpType" style="width: 125px;">
  
<%
					strSQL = " util_HTML_option_list_from_table('ha_EmployeesFormatted', 'EmpType', '" & eType & "')"
					rs.Open strSQL, connCasper10, 3, 3, 4
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
		</tr>
				
		<tr>
			<td align="right">
				Contractor:
			</td>
			<td>
				<Select id="Contractor" name="Contractor" style="width: 125px;">
  
<%
					strSQL = "util_HTML_option_list('ha_types_repl', 'class', 'YesNoNo', 'value', 'description', 'sequence', 'isDefault', '" & eContractor & "')"
					rs.Open strSQL, connCasper10, 3, 3, 4
					If not rs.eof Then
						While not rs.eof
							Response.write rs("result")
							rs.MoveNext
						Wend
					End If
					rs.Close
%>
		
				</Select>
				<span class="grayTag">Not sure why this question is repeated?</span>
			</td>
		</tr>
		
		<tr>
			<td align="right">
				Project Specialist:
			</td>
			<td>
				<Select id="ProjectSpecialist" name="ProjectSpecialist" style="width: 125px;">
  
<%
					strSQL = "util_HTML_option_list('ha_types_repl', 'class', 'YesNoNo', 'value', 'description', 'sequence', 'isDefault', '" & eProjectSpecialist & "')"
					rs.Open strSQL, connCasper10, 3, 3, 4
					If not rs.eof Then
						While not rs.eof
							Response.write rs("result")
							rs.MoveNext
						Wend
					End If
					rs.Close
%>
		
				</Select>
				<span class="grayTag">For "To Do"s</span>
			</td>
		</tr>
		
		<tr>
		  <td align="right">
			Time Clock:
		  </td>
		  <td>
				<Select id="TimeClock" name="TimeClock" style="width: 125px;">
  
<%
					strSQL = "util_HTML_option_list('ha_types_repl', 'class', 'YesNoNo', 'value', 'description', 'sequence', 'isDefault', '" & eTimeClock & "')"
					rs.Open strSQL, connCasper10, 3, 3, 4
					If not rs.eof Then
						While not rs.eof
							Response.write rs("result")
							rs.MoveNext
						Wend
					End If
					rs.Close
%>
		
				</Select>
				<span class="grayTag">Allows Time Entry</span>
		  </td>		 
		</tr>
		
		<tr>
		  <td align="right">
			Staff Status:
		  </td>
		  <td>
				<Select id="StaffStatus" name="StaffStatus" style="width: 125px;">
  
<%
					strSQL = "util_HTML_option_list('ha_types_repl', 'class', 'YesNoNo', 'value', 'description', 'sequence', 'isDefault', '" & eStaffStatus & "')"
					rs.Open strSQL, connCasper10, 3, 3, 4
					If not rs.eof Then
						While not rs.eof
							Response.write rs("result")
							rs.MoveNext
						Wend
					End If
					rs.Close
%>
		
				</Select>
				<span class="grayTag">Allows "Staff Status" Entry</span>
		  </td>		 
		</tr>
		
		<tr>
		  <td align="right">
			Termination Date:
		  </td>
		  <td>
			<input type="text" name="TerminationDate" id="TerminationDate"
				size="10" value="<%=eTerminationDate%>">
			<%=PopupDate("TerminationDate")%>
		  </td>		 
		</tr>
		
	  </table>
    </td>
  </tr>
</table>

    <span class="popupCredit"><%=CreditLine(eCreatedBy, eCreatedDate, eEditedBy, eEditedDate)%></span>	

    <span class="popupButtons">
      <a href="javascript:document.frm.submit();" class="bo_Button80">Submit</a>
      <a href="javascript:window.close();" class="bo_Button80">Close</a>
    </span>

  </form>
</body>
</html>
