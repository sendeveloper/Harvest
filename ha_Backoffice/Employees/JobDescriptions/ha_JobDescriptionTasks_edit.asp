<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsHelp.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->
  
<%
	Dim rs: Set rs = Server.CreateObject("ADODB.Recordset")

	If IsNull(Request("ID")) or Request("ID")=0 Then
		ID = 0
		Title = "Job Description Tasks Add"

		eDate = ""
		eDescription = ""
	Else
		ID = Request("ID")
		Title = "Job Description Tasks Edit"
		
		strSQL = "SELECT * FROM ha_JobDescriptionTasks WHERE ID = " & ID
        rs.open strSQL, connCasper10, 0, 1

		eCategory 		= Null2Space(rs("Category"))
		eTaskName 		= Null2Space(rs("TaskName"))
		eTaskDescription= Null2Space(rs("TaskDescription"))
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
		width: 55em;   
		}
	
    select
		{
		width: 12.5em;   
		}
	
    textarea 
		{
		font-family:    Verdana, Arial, Helvetica, sans-serif;
		width: 			55em;   
		height: 		28em; 
		resize: 		none; 
		}
  </style>
</head>

<body onLoad="SetScreen(1100,700);">
  <form method="Post" action="ha_JobDescriptionTasks_post.asp?ID=<%=ID%>" name="frm">

	  <span class="popupHeading"><%=Title%></span>
	  
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">	  
  <tr>
    <td colspan="2">
	
      <table width="95%" border="0" cellspacing="2" cellpadding="2" align="center">

		<tr><td>&nbsp;</td></tr>
			
		<tr>
		  <td width="15%" align="right">
			Category:
		  </td>
		  <td width="85%">
			<Select id="Category" name="Category">
  
<%
					strSQL = " util_HTML_option_list('ha_types_repl', 'class', 'JobDescriptions - TaskCategories', 'value', 'value', 'sequence', '', '" & eCategory & "')"
					rs.Open strSQL, connCasper10, 3, 3, 4
					If not rs.eof Then
						While not rs.eof
							Response.write Replace(rs("result"), "(blank)", "")
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
			Task Name:
		  </td>
		  <td>
			<input type="text" name="TaskName" id="TaskName"
				size="65" value="<%=eTaskName%>">				
		  </td>		 
		</tr>

		<tr vAlign="top">
			<td align="right">
				Task Description:
			</td>
			<td>
				<textarea id="TaskDescription" name="TaskDescription"><%=eTaskDescription%></textarea>
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
