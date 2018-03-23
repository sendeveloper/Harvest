<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsDates.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsHelp.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->
  <!--#include virtual="ha_backoffice/includes/Dates/ConvDate.inc"-->
  
<%
	Dim rs: Set rs = Server.CreateObject("ADODB.Recordset")

	If IsNull(Request("ID")) or Request("ID")=0 Then
		ID = 0
		Title = "Password Add"

        eID 			= ""
        eCompanyName 	= ""
        eURL 			= ""
        ePrimaryEmail 	= ""
        eBackupEmails 	= ""
        eUsername 		= ""
        ePassword 		= ""
        eCodeNumbers 	= ""
        ePurpose 		= ""
        eClass 			= ""
        ePermissionLevel = ""
		eCreatedBy		= ""
		eCreatedDate	= ""
		eEditedBy		= ""
		eEditedDate		= ""
	Else
		ID = Request("ID")
		Title = "Password Edit"
		
		strSQL = "SELECT * FROM ha_Passwords WHERE ID = " & ID
        rs.open strSQL, connCasper10, 0, 1
		
        eID 			= rs("ID")
        eCompanyName 	= rs("CompanyName")
        eURL 			= rs("URL")
        ePrimaryEmail 	= rs("PrimaryEmail")
        eBackupEmails 	= rs("BackupEmails")
        eUsername 		= rs("Username")
        ePassword 		= rs("Password")
        eCodeNumbers 	= rs("CodeNumber")
        ePurpose 		= rs("Purpose")
        eClass 			= rs("Class")
        ePermissionLevel = rs("PermissionLevel")
		eCreatedBy		= Null2Space(rs("CreatedBy"))
		eCreatedDate	= Null2Space(rs("CreatedDate"))
		eEditedBy		= Null2Space(rs("EditedBy"))
		eEditedDate		= Null2Space(rs("EditedDate"))

		
        rs.close
	End If
	
%>

<html>
<head>
  <title><%=Title%></title>

  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
  
  <link href="/ha_Backoffice/includes/ha_BackOfficePopup.css" type="text/css" rel="stylesheet">
  <script language="javascript" type="text/javascript" src="/includes/haFunctions.js"></script>
</head>

<body onLoad="SetScreen(650,600);">
  <form method="Post" action="ha_PasswordManagement_post.asp?ID=<%=ID%>" name="frm">

	  <span class="popupHeading"><%=Title%></span>
	  
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">	  
  <tr>
    <td colspan="2">
	
      <table width="95%" border="0" cellspacing="2" cellpadding="2" align="center">

		<tr><td>&nbsp;</td></tr>
			
		<tr>
		  <td width="25%" align="right">
			ID:
		  </td>
		  <td width="75%">
		    <%=eID%>
		  </td>		 
		</tr>

		<tr>
		  <td align="right">
			Company Name:
		  </td>
		  <td>
			<input type="text" name="CompanyName" id="CompanyName"
				size="40" value="<%=eCompanyName%>">				
		  </td>		 
		</tr>

		<tr>
		  <td align="right">
			URL:
		  </td>
		  <td>
			<input type="text" name="URL" id="URL"
				size="40" value="<%=eURL%>">				
		  </td>		 
		</tr>

		<tr>
		  <td align="right">
			Primary E-mail:
		  </td>
		  <td>
			<input type="text" name="PrimaryEmail" id="PrimaryEmail"
				size="40" value="<%=ePrimaryEmail%>">				
		  </td>		 
		</tr>

		<tr>
		  <td align="right">
			Backup E-mails:
		  </td>
		  <td>
			<input type="text" name="BackupEmails" id="BackupEmails"
				size="40" value="<%=eBackupEmails%>">				
		  </td>		 
		</tr>

		<tr>
		  <td align="right">
			Username:
		  </td>
		  <td>
			<input type="text" name="Username" id="Username"
				size="40" value="<%=eUsername%>">				
		  </td>		 
		</tr>

		<tr>
		  <td align="right">
			Password:
		  </td>
		  <td>
			<input type="text" name="Password" id="Password"
				size="40" value="<%=ePassword%>">				
		  </td>		 
		</tr>

		<tr>
		  <td align="right">
			Code Numbers:
		  </td>
		  <td>
			<input type="text" name="CodeNumbers" id="CodeNumbers"
				size="40" value="<%=eCodeNumbers%>">				
		  </td>		 
		</tr>

		<tr>
		  <td align="right">
			Purpose:
		  </td>
		  <td>
			<input type="text" name="Purpose" id="Purpose"
				size="40" value="<%=ePurpose%>">				
		  </td>		 
		</tr>

		<tr>
			<td align="right">
				Class:
			</td>
			<td>
				<Select id="Class" name="Class" style="width: 275px;">
  
<%
					strSQL = " util_HTML_option_list('ha_types_repl', 'class', 'Passwords - Class', 'value', 'description', 'sequence', '', '" & eClass & "')"
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
				Permission:
			</td>
			<td>
				<Select id="Permission" name="Permission" style="width: 275px;">
  
<%
					strSQL = " util_HTML_option_list('ha_types_repl', 'class', 'Passwords - Permission', 'value', 'description', 'sequence', '', '" & ePermissionLevel & "')"
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

		<tr><td>&nbsp;</td></tr>
		
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
