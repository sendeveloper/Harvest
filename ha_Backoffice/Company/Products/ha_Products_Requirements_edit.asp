<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsHelp.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->
  
<%
  	Set connLocal = Server.CreateObject("ADODB.Connection")
	connLocal.Open "driver=SQL Server; server=127.0.0.1,7043; uid=davewj2o; pwd=get2it; database=ha_PublishedTables"
	Dim rs: Set rs = Server.CreateObject("ADODB.Recordset")
  
	If IsNull(Request("ID")) or Request("ID")=0 Then
	Else
		ID = Request("ID")

		strSQL = "SELECT * FROM ha_Products WHERE ID = " & ID
		rs.Open strSQL, connLocal, 3, 3

		If Not rs.EOF Then
			eItemID							= rs("ItemID")
			eDescription 					= rs("Description")
			
			eRequiresShipping 				= rs("RequiresShipping")
			eRequiresRegistration_rt 		= rs("RequiresRegistration_rt")
			eRequiresRegistration_nm 		= rs("RequiresRegistration_nm")
			eRequiresRegistration_nmpro 	= rs("RequiresRegistration_nmpro")
			eRequiresRegistration_lookup 	= rs("RequiresRegistration_lookup")
			eRequiresEmailTables 			= rs("RequiresEmailTables")
			eRequiresRegistration_initial 	= rs("RequiresRegistration_initial")
			eRequiresRegistration_updates 	= rs("RequiresRegistration_updates")
			eRequiresRegistration_link 		= rs("RequiresRegistration_link")
			eRequiresProcedureSetup 		= rs("RequiresProcedureSetup")
			
			eCreatedBy		= Null2Space(rs("CreatedBy"))
			eCreatedDate	= Null2Space(rs("CreatedDate"))
			eEditedBy		= Null2Space(rs("EditedBy"))
			eEditedDate		= Null2Space(rs("EditedDate"))
		End If
	End If
	
	rs.close
%>

<html>
<head>
  <title>Harvest American Backoffice - Edit Product Requirements</title>
  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
  
  <link href="/ha_BackOffice/includes/ha_BackofficePopup.css" type="text/css" rel="stylesheet" media="screen" />
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
  </style>
</head>

<body onload="SetScreen(850,650)">
  <form action="ha_Products_Requirements_post.asp?ID=<%=ID%>" method="post" name="frm">

	<span class="popupHeading">Product Requirements Edit</span>

    <table width="95%" border="0" cellspacing="2" cellpadding="2">

		<tr><td>&nbsp;</td></tr>
		
		<tr>
		  <td width="40%" align="right">
			Item ID:
		  </td>
		  <td width="60%" style="color: #1F3A4A; font-weight: bold;">
			<%=eItemID%>
		  </td>		 
		</tr>
		
		<tr>
		  <td align="right">
			Description:
		  </td>
		  <td style="color: #1F3A4A; font-weight: bold;">
			<%=eDescription%>
		  </td>		 
		</tr>
		
		<tr><td>&nbsp;</td></tr>
		
		<%=YNDropdown ("Shipping", "RequiresShipping", eRequiresShipping, 19)%>
		<%=YNDropdown ("Registration Raffle Ticket", "RequiresRegistration_rt", eRequiresRegistration_rt, 20)%>
		<%=YNDropdown ("Registration Number Machine", "RequiresRegistration_nm", eRequiresRegistration_nm, 21)%>
		<%=YNDropdown ("Registration Number Machine Pro", "RequiresRegistration_nmpro", eRequiresRegistration_nmpro, 22)%>
		<%=YNDropdown ("Registration Lookup", "RequiresRegistration_lookup", eRequiresRegistration_lookup, 23)%>
		<%=YNDropdown ("E-mail of Tables", "RequiresEmailTables", eRequiresEmailTables, 24)%>
		<%=YNDropdown ("Registration Initial", "RequiresRegistration_initial", eRequiresRegistration_initial, 25)%>
		<%=YNDropdown ("Registration Updates", "RequiresRegistration_updates", eRequiresRegistration_updates, 26)%>
		<%=YNDropdown ("Registration Link", "RequiresRegistration_link", eRequiresRegistration_link, 27)%>
		<%=YNDropdown ("Procedure Set-up", "RequiresProcedureSetup", eRequiresProcedureSetup, 28)%>
		
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

<%
Function YNDropdown (Title, ID, Value, HelpID)

	s = "<tr>"
	s = s & "  <td align=""right"">"
	s = s & "    " & Title & ":"
	s = s & "  </td>"
	s = s & "  <td>"
	s = s & "    <Select id=""" & ID & """ name=""" & ID & """>"
	
				strSQL = "ha_BackOffice.dbo.util_HTML_option_list('ha_PublishedTables.dbo.ha_types', 'class', 'TrueFalse', 'value', 'description', 'sequence', '', '" & Value & "')"
			
				rs.Open strSQL, connLocal, 3, 3, 4
				If not rs.eof Then
					While not rs.eof
						s = s & rs("result")
						rs.MoveNext
					Wend
				End If
				rs.Close	
	
	s = s & "    </Select>"
	s = s & "    " & PopupHelp(HelpID)
	s = s & "  </td>"
	s = s & "</tr>"
	
	YNDropdown = s

End Function
%>
