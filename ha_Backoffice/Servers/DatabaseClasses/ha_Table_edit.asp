<!DOCTYPE html>
<!--#include virtual="includes/connection.asp"-->
<!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->
  
<%
	Dim rs: Set rs = Server.CreateObject("ADODB.Recordset")

	If IsNull(Request("ID")) or Request("ID")=0 Then
		ID = 0
		Title = "Table Add"

		eID 			= ""
		eClassName 		= ""
		eTableName 		= ""
		eUpdateMethod 	= ""
		ePurpose		= ""
	Else
		ID = Request("ID")
		Title = "Table Edit"
		
		strSQL = "SELECT * FROM ha_Database_Objects WHERE ID = " & ID
        rs.open strSQL, connCasper10, 0, 1

		eID 			= Null2Space(rs("ID"))
		eClassName		= Null2Space(rs("ClassName"))		
		eTableName 		= Null2Space(rs("TableName"))
		eUpdateMethod	= Null2Space(rs("UpdateMethod"))
		ePurpose		= Null2Space(rs("Purpose"))
		eCreatedBy		= Null2Space(rs("CreatedBy"))
		eCreatedDate	= Null2Space(rs("CreatedDate"))
		eEditedBy		= Null2Space(rs("EditedBy"))
		eEditedDate		= Null2Space(rs("EditedDate"))
		
        rs.close
	End If
	
%>

<html>
<head>
  <title>Harvest American Backoffice - <%=Title%></title>
  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
  
  <link href="/ha_Backoffice/includes/ha_BackOfficePopup.css" type="text/css" rel="stylesheet">
  <script language="javascript" type="text/javascript" src="/includes/haFunctions.js"></script>
  
  <style type="text/css">  
    input[type="text"]
		{
		font-family: Verdana, Arial, Helvetica, sans-serif;
		width: 30em;   
		}
		
	textarea.Purpose
		{
		font-family: Verdana, Arial, Helvetica, sans-serif;
		width: 30em;   
		height: 8em; 
		resize: none; 
		}
	</style>
		
</head>

<body onLoad="SetScreen(700,500);">
  <form method="Post" action="ha_Table_post.asp?ID=<%=ID%>" name="frm">

	  <span class="popupHeading"><%=Title%></span>
	  
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">	  
  <tr>
    <td colspan="2">
	
      <table width="95%" border="0" cellspacing="2" cellpadding="2" align="center">

		<tr><td>&nbsp;</td></tr>
			
		<tr>
		  <td width="20%" align="right">
			ID:
		  </td>
		  <td width="80%">
			<%=eID%>
		  </td>		 
		</tr>
		
		<tr>
		  <td align="right">
			Class Name:
		  </td>
		  <td>
			<input type="text" name="ClassName" id="ClassName"
				size="50" value="<%=eClassName%>">
		  </td>		 
		</tr>
		
		<tr>
		  <td align="right">
			Table Name:
		  </td>
		  <td>
			<input type="text" name="TableName" id="TableName"
				size="50" value="<%=eTableName%>">
		  </td>		 
		</tr>

		<tr>
		  <td align="right">
			Update Method:
		  </td>
		  <td>
			<input type="text" name="UpdateMethod" id="UpdateMethod"
				size="50" value="<%=eUpdateMethod%>">				
		  </td>		 
		</tr>

		<tr>
		  <td align="right" vAlign="top">
			Purpose:
		  </td>
		  <td>
			<textarea cols="1" rows="1" id="Purpose" name="Purpose" class="Purpose"><%=ePurpose%></textarea>				
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
