<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
 <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/Config.asp"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsHelp.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->
  
<%
  	Set connLocal = Server.CreateObject("ADODB.Connection")
	connLocal.Open "driver=SQL Server; server=127.0.0.1,7043; uid=davewj2o; pwd=get2it; database=ha_PublishedTables"
	Dim rs: Set rs = Server.CreateObject("ADODB.Recordset")

	If IsNull(Request("ID")) or Request("ID")=0 or Request("ID")="" Then
		ID = 0
		Title = "Server Maintenance Log Add"
		
		eInventoried = 0
	Else
		ID = Request("ID")
		Title = "Server Maintenance Log Edit"
		
		strSQL = "SELECT * FROM ha_Server_Maintenance_Log WHERE ID = " & ID
        rs.open strSQL, connCasper10, 0, 1

	    If Not rs.EOF Then
			eID					= rs("ID")
			eServerName 		= rs("ServerName")
			
			eTask		 		= rs("Task")
			eNotes 				= rs("Notes")
			eEngineer		 	= rs("Engineer")
			eType 				= rs("Type")
					
			eCreatedBy			= Null2Space(rs("CreatedBy"))
			eCreatedDate		= Null2Space(rs("CreatedDate"))
			eEditedBy			= Null2Space(rs("EditedBy"))
			eEditedDate			= Null2Space(rs("EditedDate"))
		End If
		
        rs.close
	End If
%>

<html>
<head>

  <title>Harvest American Backoffice - <%=Title%></title>
  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
  
  <link href="/ha_BackOffice/includes/ha_BackofficePopup.css" type="text/css" rel="stylesheet">
  <script type="text/javaScript" src="/ha_BackOffice/includes/haFunctions.js"></script> 
  
  <style type="text/css">    
    input[type="text"]
		{
		width: 12em;   
		}
				
	select
		{
		width: 12.4em;   
		}
	
	textarea
		{
		font-family:Verdana, Arial, Helvetica, sans-serif;
		font-size: 	10pt;
		width: 		400px;
		height:		3em;
		resize: 	none;
		}
  </style>
</head>

<body onload="SetScreen(850,850)">
  <form action="ha_ServerMaintenance_post.asp?ID=<%=eID%>" method="post" name="frm">

	<span class="popupHeading"><%=Title%></span>

    <table width="95%" border="0" cellspacing="2" cellpadding="2">

		<tr><td>&nbsp;</td></tr>
		
		<tr>
		  <td width="30%" align="right">
			ID:
		  </td>
		  <td width="70%" style="color: #1F3A4A; font-weight: bold;">
			<%=eID%>
		  </td>		 
		</tr>
		
		<tr>
		  <td align="right">
			Server Name:
		  </td>
		  <td style="color: #1F3A4A; font-weight: bold;">
		  
			<Select id="ServerName" name="ServerName">
				
<%
				strSQL = "ha_BackOffice.dbo.util_HTML_option_list('ha_types_repl', 'description', 'Server Name', 'Value', 'Value', 'sequence', '', '" & eServerName & "')"
		
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
			<%=PopupHelp(34)%>
		  </td>		 
		</tr>
<tr>
		  <td align="right">
			Type:
		  </td>
		  <td>
			<Select id="Type" name="Type">
				
<%
				strSQL = "ha_BackOffice.dbo.util_HTML_option_list('ha_PublishedTables.dbo.ha_types', 'class', 'ServerMaintenanceType', 'value', 'description', 'sequence', '', '" & eType & "')"
			
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
			<%=PopupHelp(37)%>
		  </td>		 
		</tr>
					
		<tr>
		  <td align="right"  style="vertical-align:top;">
			Task:
		  </td>
		  <td>
			<textarea name="Task" id="Task" style="height:150px;"><%=eTask%></textarea>
			<%=PopupHelp(35)%>
		  </td>		 
		</tr>

		<tr>
		  <td align="right"  style="vertical-align:top;">
			Notes:
		  </td>
		  <td>
		  <textarea name="Notes" id="Notes"  style="height:150px;" ><%=eNotes%></textarea>
						
				
		  </td>		 
		</tr>
		
		
		<tr>
		  <td align="right">
			Engineer:
		  </td>
		  <td>
			<input type="text" name="Engineer" id="Engineer"
				size="10" value="<%=eEngineer%>">
			<%=PopupHelp(36)%>
		  </td>		 
		</tr>

		<tr><td>&nbsp;</td></tr>
		
    </table>
		
		
    <span class="popupCredit"><%=CreditLine(eCreatedBy, eCreatedDate, eEditedBy, eEditedDate)%></span>	

    <span class="popupButtons">
      <a href="javascript:document.frm.submit();" class="bo_Button80">Submit</a>
      <a href="javascript:window.close();" class="bo_Button80">Close</a>
    </span>	
	
  </form>
</body>
</html>
