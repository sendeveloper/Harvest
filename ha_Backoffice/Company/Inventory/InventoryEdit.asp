<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsDates.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsHelp.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->
  
<%
	Dim rs: Set rs = Server.CreateObject("ADODB.Recordset")

	If IsNull(Request("ID")) or Request("ID")=0 Then
		ID = 0
		Title = "Inventory Add"

		eType			= ""
		eDescription	= ""
		eCondition		= ""
		eAssignedTo		= ""
		eQuantity		= ""

	Else
		ID = Request("ID")
		Title = "Inventory Edit"
		
		strSQL = "SELECT * FROM ha_Inventory WHERE ID = " & ID
        rs.open strSQL, connCasper10, 0, 1
		
		eType			= Null2Space(rs("Type"))
		eDescription	= Null2Space(rs("Description"))
		eCondition		= Null2Space(rs("Condition"))
		eAssignedTo		= Null2Space(rs("AssignedTo"))
		eQuantity		= Null2Space(rs("Quantity"))
		
        rs.close
	End If
		
%>

<html>
<head>
  <title>Add Inventory</title>
  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
  
  <link href="/ha_Backoffice/includes/ha_BackOfficePopup.css" type="text/css" rel="stylesheet">
  <script language="javascript" type="text/javascript" src="/includes/haFunctions.js"></script>
</head>

<body onLoad="SetScreen(700, 400);">
  <form method="Post" action="BlankPost.asp?ID=<%=ID%>" name="frm">

	  <span class="popupHeading"><%=Title%></span>
	  
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">	  
  <tr>
    <td colspan="2">
	
      <table width="95%" border="0" cellspacing="2" cellpadding="2" align="center">

		<tr><td>&nbsp;</td></tr>
			
		<tr>
		  <td width="35%" align="right">
			Type:
		  </td>
		  <td width="65%">
			<input type="text" name="Type" id="Type"
				size="20" value="<%=eType%>">
		  </td>		 
		</tr>

		<tr>
		  <td align="right">
			Description:
		  </td>
		  <td>
			<input type="text" name="Description" id="Description"
				size="20" value="<%=eDescription%>">
		  </td>		 
		</tr>
		
		<tr>
		  <td align="right">
			Condition:
		  </td>
		  <td>
			<input type="text" name="Condition" id="Condition"
				size="20" value="<%=eCondition%>">
		  </td>		 
		</tr>
		
		<tr>
		  <td align="right">
			Assigned To:
		  </td>
		  <td>
			<input type="text" name="AssignedTo" id="AssignedTo"
				size="20" value="<%=eAssignedTo%>">
		  </td>		 
		</tr>
		
		<tr>
		  <td align="right">
			Quantity:
		  </td>
		  <td>
			<input type="text" name="Quantity" id="Quantity"
				size="20" value="<%=eQuantity%>">
		  </td>		 
		</tr>
		
		<tr><td>&nbsp;</td></tr>
		
	  </table>
    </td>
  </tr>
</table>
	
    <span class="popupButtons">
      <a href="javascript:document.frm.submit();" class="bo_Button80">Submit</a>
      <a href="javascript:window.close();" class="bo_Button80">Close</a>
    </span>

  </form>
</body>
</html>
