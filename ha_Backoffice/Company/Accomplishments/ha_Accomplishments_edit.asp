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
		Title = "Accomplishment Add"

		eDate = ""
		eDescription = ""
	Else
		ID = Request("ID")
		Title = "Accomplishment Edit"
		
		strSQL = "SELECT * FROM ha_Accomplishments WHERE ID = " & ID
        rs.open strSQL, connCasper10, 0, 1

		eDate 			= Null2Space(rs("Date"))
		eDescription 	= Null2Space(rs("Description"))
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
</head>

<body onLoad="SetScreen(700,400);">
  <form method="Post" action="ha_Accomplishments_post.asp?ID=<%=ID%>" name="frm">

	  <span class="popupHeading"><%=Title%></span>
	  
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">	  
  <tr>
    <td colspan="2">
	
      <table width="95%" border="0" cellspacing="2" cellpadding="2" align="center">

		<tr><td>&nbsp;</td></tr>
			
		<tr>
		  <td width="15%" align="right">
			Date:
		  </td>
		  <td width="85%">
			<input type="text" name="AccomplishmentDate" id="AccomplishmentDate"
				size="10" value="<%=eDate%>">
			<%=PopupDate("AccomplishmentDate")%>
			<%=PopupHelp(16)%>
		  </td>		 
		</tr>

		<tr>
		  <td align="right">
			Description:
		  </td>
		  <td>
			<input type="text" name="Description" id="Description"
				size="65" value="<%=eDescription%>">				
			<%=PopupHelp(17)%>
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
