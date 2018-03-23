<!--#include virtual="includes/connection.asp"-->

<html>
<head>
    <title>Job Description Titles Post</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />

</head>
<body>

<%
    Dim rs, strSQL

    Set rs = Server.CreateObject("ADODB.Recordset")
	
	If isnull(Request("TitleID")) or Request("TitleID") = "" Then
		TitleID = 0
	Else
		TitleID = Request("TitleID")
	End If
	
	'See if we need a new title and get the ID first
	If TitleID = 0 Then
		strSQL = "ha_Table_Row_create('ha_JobDescriptionTitles', '" & Session("ha_shortname") & "')"
		rs.Open strSQL, connLocal, 3, 3, 4
			If not rs.eof Then
				TitleID = rs("NewID")
			End If
		rs.Close
	End If
	
	'Save the basic data to the title table
	strSQL = "SELECT * FROM ha_JobDescriptionTitles WHERE ID = " & cstr(TitleID)
	rs.Open strSQL, connLocal, 2, 3
		rs("Category") 		= Request("Category")
		rs("TitleName") 	= ltrim(rtrim(Request("TitleName")))
		rs("TitlePurpose") 	= ltrim(rtrim(Request("Purpose")))
		
		rs("EditedBy") 		= Session("ha_shortname")
		rs("EditedDate") 	= Now
	rs.Update
	rs.Close
	
	Select Case Request("SaveAction")
	Case "AssigneeAttach"
		'PassID = EmployeeID
		strSQL = "ha_JobDescriptionEmployee_attach(" & TitleID & ", " & Request("PassID") & ", '" & Session("ha_shortname") & "')"
		connLocal.Execute strSQL
	Case "AssigneeRemove"
		'PassID = Job Title Links's Table ID;
		strSQL = "UPDATE ha_JobDescriptionTitles_links SET DeletedBy = '" & Session("ha_shortname") & "', DeletedDate = GETDATE() WHERE ID = " & Request("PassID")
		connLocal.Execute strSQL
	Case "TaskMoveDown"
		'PassID = Job Task Links's Table ID;
		strSQL = "ha_JobDescriptionTask_move('Down', " & Request("PassID") & ", '" & Session("ha_shortname") & "')"
		connLocal.Execute strSQL
	Case "TaskMoveUp"
		'PassID = Job Task Links's Table ID;
		strSQL = "ha_JobDescriptionTask_move('Up', " & Request("PassID") & ", '" & Session("ha_shortname") & "')"
		connLocal.Execute strSQL
	Case "TaskAttach"
		'PassID = TaskID
		strSQL = "ha_JobDescriptionTask_attach(" & Request("PassID") & ", " & TitleID & ", '" & Session("ha_shortname") & "')"
		connLocal.Execute strSQL
	Case "TaskRemove"
		'PassID = Job Task Links's Table ID;
		strSQL = "UPDATE ha_JobDescriptionTasks_links SET DeletedBy = '" & Session("ha_shortname") & "', DeletedDate = GETDATE() WHERE ID = " & Request("PassID")
		connLocal.Execute strSQL
	Case "SaveFields"
		CloseIt = True
	End Select
%>

<script type="text/javascript">
    window.opener.location.href = window.opener.location.href;
<%
	If CloseIt = True Then
%>
		window.close();
<%
	Else
%>

		//Returns you to previous page and refreshes 
		var backLocation = document.referrer;
		if (backLocation) 
			{
			if (backLocation.indexOf("?") > -1) 
				{
				backLocation += "&randomParam=" + new Date().getTime();
				} 
			else 
				{
				backLocation += "?randomParam=" + new Date().getTime();
				}
			window.location.assign(backLocation);
			}
<%
	End If
%>
</script>

</body>	
</html>
