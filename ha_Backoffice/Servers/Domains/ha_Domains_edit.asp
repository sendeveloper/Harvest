<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsHelp.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->
  <!--#include virtual="ha_backoffice/includes/Dates/ConvDate.inc"-->
  
<%
	Dim rs: Set rs = Server.CreateObject("ADODB.Recordset")

	If IsNull(Request("ID")) or Request("ID")=0 Then
		ID = 0
		Title = "Domain Add"

		eSubName 			= ""
		eDomainName 		= ""
		ePurpose 			= ""
		eClass 				= "Harvest"
		eServerName 		= ""
		eSecureSite 		= ""
		eFTPLogin 			= ""
		eFTPPassword 		= ""
		eDomainRegister 	= "GoDaddy"
		eDB 				= ""
		eNotes 				= ""
		eCreatedBy			= ""
		eCreatedDate		= ""
		eEditedBy			= ""
		eEditedDate			= ""

	Else
		ID = Request("ID")
		Title = "Domain Edit"
		
		strSQL = "SELECT * FROM ha_Domains WHERE ID = " & ID
        rs.open strSQL, connCasper10, 0, 1

		eSubName			= Null2Space(rs("SubName"))
		eDomainName 		= Null2Space(rs("DomainName"))
		ePurpose 			= Null2Space(rs("Purpose"))
		eClass 				= Null2Space(rs("Class"))
		eServerName 		= Null2Space(rs("ServerName"))
		eSecureSite 		= Null2Space(rs("SecureSite"))
		eFTPLogin 			= Null2Space(rs("FTPLogin"))
		eFTPPassword 		= Null2Space(rs("FTPPassword"))
		eDomainRegister 	= Null2Space(rs("DomainRegister"))
		eDB 				= Null2Space(rs("DB"))
		eNotes 				= Null2Space(rs("Notes"))
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
</head>

<body onLoad="SetScreen(650,650);">
  <form method="Post" action="ha_Domains_post.asp?ID=<%=ID%>" name="frm">

	  <span class="popupHeading"><%=Title%></span>
	  
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">	  
  <tr>
    <td colspan="2">
	
      <table width="95%" border="0" cellspacing="2" cellpadding="1" align="center">

		<tr><td>&nbsp;</td></tr>
			
		<tr>
		  <td width="25%" align="right">
			SubDomain:
		  </td>
		  <td width="75%">
			<input type="text" name="SubName" id="SubName"
				size="15" value="<%=eSubName%>">
			<span style="color: #C0C0C0;">(Leave Blank for www, no dot)</span>
		  </td>		 
		</tr>

		<tr>
		  <td align="right">
			Domain Name:
		  </td>
		  <td>
			<input type="text" name="DomainName" id="DomainName"
				size="40" value="<%=eDomainName%>">				
		  </td>		 
		</tr>

		<tr>
		  <td align="right" valign="top">
			Purpose:
		  </td>
		  <td>
			<textarea name="Purpose" id="Purpose" cols="40" rows="3"><%=ePurpose%></textarea>
		  </td>		 
		</tr>
		
		<tr>
			<td align="right">
				Class:
			</td>
			<td>
				<Select id="Class" name="Class" style="width: 125px;">
  
<%
					strSQL = " util_HTML_option_list_from_table('ha_Domains', 'Class', '" & eClass & "')"
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
				Server Name:
			</td>
			<td>
				<Select id="ServerName" name="ServerName" style="width: 125px;">
  
<%
					strSQL = " util_HTML_option_list_from_table('ha_Domains', 'ServerName', '" & eServerName & "')"
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
				Domain Registrar:
			</td>
			<td>
				<Select id="DomainRegister" name="DomainRegister" style="width: 125px;">
  
<%
					strSQL = " util_HTML_option_list_from_table('ha_Domains', 'DomainRegister', '" & eDomainRegister & "')"
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
				Secure Website:
			</td>
			<td>
				<Select id="SecureSite" name="SecureSite" style="width: 125px;">
  
<%
					strSQL = "util_HTML_option_list('ha_types_repl', 'class', 'YesNoNo', 'value', 'description', 'sequence', 'isDefault', '" & eSecureSite & "')"
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
			FTP Login:
		  </td>
		  <td>
			<input type="text" name="FTPLogin" id="FTPLogin"
				size="15" value="<%=eFTPLogin%>">				
		  </td>		 
		</tr>
		
		<tr>
		  <td align="right">
			FTP Password:
		  </td>
		  <td>
			<input type="text" name="FTPPassword" id="FTPPassword"
				size="15" value="<%=eFTPPassword%>">				
		  </td>		 
		</tr>
		
		<tr>
		  <td align="right">
			Main Database:
		  </td>
		  <td>
			<input type="text" name="DB" id="DB"
				size="40" value="<%=eDB%>">				
		  </td>		 
		</tr>
		
		<tr>
		  <td align="right" valign="top">
			Notes:
		  </td>
		  <td>
			<textarea name="Notes" id="Notes" cols="40" rows="3"><%=eNotes%></textarea>
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
