<!DOCTYPE html>
  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsDates.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsHelp.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsPermissions.inc"-->
  <!--#include virtual="ha_backoffice/includes/Dates/ConvDate.inc"-->
  
<%
	Dim rs: Set rs = Server.CreateObject("ADODB.Recordset")

	If IsNull(Request("ID")) or Request("ID")=0 Then
		ID = 0
		Title = "Calendar Add"

		eType				= ""
		eRequestFrom		= ""
		eRequestTo			= ""
		eReason				= ""
		eDepartureTime		= ""
		eRequestFor			= Session("ha_ShortName")
		eApproved			= ""
		
		eCreatedBy			= ""
		eCreatedDate		= ""
		eEditedBy			= ""
		eEditedDate			= ""		
	Else
		ID = Request("ID")
		Title = "Calendar Edit"
		
		strSQL = "SELECT * FROM ha_EmpCalendar WHERE ID = " & ID
        rs.open strSQL, connCasper10, 0, 1
		
		eType				= Null2Space(rs("Type"))
		eRequestFrom		= Null2Space(rs("RequestFrom"))
		eRequestTo			= Null2Space(rs("RequestTo"))
		eReason				= Null2Space(rs("Reason"))
		eDepartureTime		= Null2Space(rs("DepartureTime"))
		eRequestFor			= Null2Space(rs("RequestFor"))
		eDuration			= Null2Space(rs("Duration"))
		eApproved			= IIf(rs("IsApproved")=True,"Yes", "No")
				
		eCreatedBy			= Null2Space(rs("CreatedBy"))
		eCreatedDate		= Null2Space(rs("CreatedDate"))
		eEditedBy			= Null2Space(rs("EditedBy"))
		eEditedDate			= Null2Space(rs("EditedDate"))
		
        rs.close
	End If
%>

<html>
<head>
  <title><%=Title%></title>

  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
  
  <link href="/ha_Backoffice/includes/ha_BackOfficePopup.css" type="text/css" rel="stylesheet">
  <script language="javascript" type="text/javascript" src="/includes/haFunctions.js"></script>
  
  <script type="text/javascript">
  
    function clickChange() 
		{
		var e = document.getElementById("StaffSelect");
		var strUser = e.options[e.selectedIndex].value;		
		document.getElementById('RequestFor').value = strUser;
        //alert('Change name, Dave');
        }		
		
  </script>
</head>

<body onLoad="SetScreen(750,650);">
  <form method="Post" action="ha_EmployeeCalendar_post.asp?ID=<%=ID%>" name="frm">

	  <span class="popupHeading"><%=Title%></span>
	  
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">	  
  <tr>
    <td colspan="2">
	
      <table width="95%" border="0" cellspacing="2" cellpadding="1" align="center">

		<tr><td>&nbsp;</td></tr>
			
		<tr>
		  <td width="30%" align="right">
			Type:
		  </td>
		  <td width="25%">
			<Select id="Type" name="Type" style="width: 125px;">
  
<%
					strSQL = "util_HTML_option_list('ha_types_repl', 'class', 'EmployeeCalendarTypes', 'value', 'description', 'sequence', 'isDefault', '" & eType & "')"
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
		  <td width="45%">
		    &nbsp;
		  </td>
		</tr>

		<tr>
		  <td align="right">
			Request From:
		  </td>
		  <td colspan="2">
			<input type="text" name="RequestFrom" id="RequestFrom"
				size="10" value="<%=eRequestFrom%>">
				<%=PopupDate("RequestFrom")%>
		  </td>		 
		</tr>
		
		<tr>
		  <td align="right">
			Request To:
		  </td>
		  <td colspan="2">
			<input type="text" name="RequestTo" id="RequestTo"
				size="10" value="<%=eRequestTo%>">
				<%=PopupDate("RequestTo")%>
		  </td>		 
		</tr>
		
		<tr>
		  <td align="right">
			Reason:
		  </td>
		  <td colspan="2">
			<input type="text" name="Reason" id="Reason"
				size="60" value="<%=eReason%>">
		  </td>		 
		</tr>
						
		<tr>
		  <td align="right">
			Request For:
		  </td>
		  <td>
			<input type="text" name="RequestFor" id="RequestFor"
				size="13" value="<%=eRequestFor%>" readonly style="color: #999999;">
		  </td>		 
		  <td>
			<a href="javascript:clickChange();" class="bo_Button80">< Change</a>
		  
			<Select id="StaffSelect" name="StaffSelect" style="width: 150px;">
  
<%
					strSQL = "util_HTML_option_list('ha_EmployeesDropdown', 'EmpType', 'Full Time', 'ShortName', 'FullName', 'ShortName', '', '" & eRequestFor & "')"
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
		</tr>
		
		<tr>
		  <td align="right">
				Duration:
		  </td>
		  <td>
			<Select id="Duration" name="Duration" style="width: 125px;">
  
<%
					strSQL = "util_HTML_option_list('ha_types_repl', 'class', 'EmployeeCalendarDuration', 'value', 'description', 'sequence', 'isDefault', '" & eDuration & "')"
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
		</tr>
		
		<tr>
		  <td align="right">
			Departure Time:
		  </td>
		  <td>
			<input type="text" name="DepartureTime" id="DepartureTime"
				size="13" value="<%=eDepartureTime%>">
		  </td>		 
		</tr>
								
		<tr>
		  <td align="right">
				Approved:
		  </td>
		  <td>
			<Select id="Approved" name="Approved" style="width: 125px;">
  
<%
					strSQL = "util_HTML_option_list('ha_types_repl', 'class', 'YesNoNo', 'value', 'description', 'sequence', 'isDefault', '" & eApproved & "')"
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
