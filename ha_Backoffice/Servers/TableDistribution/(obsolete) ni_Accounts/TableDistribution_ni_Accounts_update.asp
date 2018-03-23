<!--#include virtual="ha_backoffice/Servers/TableDistribution/includes/TableDistribution_TaskLog_write.inc"-->
<!--#include virtual="ha_backoffice/Servers/TableDistribution/includes/TableDistribution_Functions.inc"-->

<%
    Dim rs
    Dim SQL
    Dim conn
	Dim ExecBy
	Dim AccountID
	Dim RequestServer
	Dim RequestDatabase
	Dim RequestExecBy
	Dim RequestError
	Dim RequestStart
	Dim RequestEnd
	Dim PostStatus
	Dim TaskName

	RequestError = 0
	ReadRequestVariables
	
	If isnull(Request("aID")) or Request("aID") = "" Then
		AccountID = ""
		TaskName = "TableDistribution_ni_Accounts_update (Full)"
	Else
		AccountID = Request("aID")
		TaskName = "TableDistribution_ni_Accounts_update (Partial)"
	End If
		
	If isnull(Request("Start")) or Request("Start") = ""  Then
		RequestStart = 0
	Else
		RequestStart = Request("Start")
	End If

	If isnull(Request("End")) or Request("End") = "" Then
		RequestEnd = 0
	Else
		RequestEnd = Request("End")
	End If
		
	If AccountID = "" Then
		SQL = "ni_Accounts_repl_Table_refresh(0, " & RequestStart & ", " & RequestEnd & ")"
	Else
		SQL = "ni_Accounts_repl_Table_refresh(" & AccountID & ")"
	End If

	

    response.ContentType = "text/xml"
    response.write "<?xml version='1.0' encoding='ISO-8859-1'?>"
    response.write "<accounts_subscriptions_refresh>"
	response.write "<request_variable_server>" & RequestServer & "</request_variable_server>"
	response.write "<request_variable_database>" & RequestDatabase & "</request_variable_database>"
	response.write "<request_variable_ExecBy>" & RequestExecBy & "</request_variable_ExecBy>"
	response.write "<request_variable_AccountID>" & AccountID & "</request_variable_AccountID>"
	response.write "<request_variable_start>" & RequestStart & "</request_variable_start>"
	response.write "<request_variable_end>" & RequestEnd & "</request_variable_end>"
	
	If RequestError = 1 Then
		PostStatus = "Error"
		response.write "<response_server_name></response_server_name>"
		response.write "<response_database_name></response_database_name>"
		response.write "<response_records_inserted></response_records_inserted>"
		response.write "<response_records_updated></response_records_updated>"
	Else
		PostStatus = "Finished"
		set conn=server.CreateObject("ADODB.Connection")
		
		conn.Open ReadConnection(RequestServer, RequestDatabase)

		set rs=server.createObject("ADODB.Recordset")
		response.write "<request_SQL>" & SQL & "</request_SQL>"
		rs.open SQL, conn, 3, 3, 4

		If not rs.eof then
			response.write "<response_server_name>" & rs("ServerName") & "</response_server_name>"
			response.write "<response_database_name>" & rs("DatabaseName") & "</response_database_name>"
			response.write "<response_records_inserted>" & rs("RecordsInserted") & "</response_records_inserted>"
			response.write "<response_records_updated>" & rs("RecordsUpdated") & "</response_records_updated>"
		End If
		
		rs.Close
		set rs = nothing
	End If
	
	response.write "<response_task_name>" & TaskName & "</response_task_name>"
    response.write "</accounts_subscriptions_refresh>"


	'Post to Task Log
	Task = RequestExecBy & " --> ni_Accounts Update: " & RequestServer & " " & RequestDatabase
	If AccountID = "" Then
		Task = Task & " Start=" & RequestStart & " End=" & RequestEnd
	Else
		Task = Task & " AccountID=" & AccountID
	End If
	Call TaskLogWrite("", TaskName, Task, PostStatus)

%>

