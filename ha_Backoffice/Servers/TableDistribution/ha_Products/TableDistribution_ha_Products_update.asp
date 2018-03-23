<!--#include virtual="ha_backoffice/Servers/TableDistribution/includes/TableDistribution_TaskLog_write.inc"-->
<!--#include virtual="ha_backoffice/Servers/TableDistribution/includes/TableDistribution_Functions.inc"-->

<%
    Dim rs
    Dim SQL
    Dim conn
	Dim ExecBy
	Dim RequestServer
	Dim RequestDatabase
	Dim RequestExecBy
	Dim RequestError
	Dim PostStatus
	
	RequestError = 0
	ReadRequestVariables	

	If isnull(Request("ID")) or Request("ID") = "" Then
		ID = ""
		TaskName = "TableDistribution_ha_Products_update (Full)"
	Else
		ID = Request("ID")
		TaskName = "TableDistribution_ha_Products_update (" & ID & ")"
	End If
	
    response.ContentType = "text/xml"
    response.write "<?xml version='1.0' encoding='ISO-8859-1'?>"
    response.write "<ha_Products_update>"
	response.write "<request_variable_server>" & RequestServer & "</request_variable_server>"
	response.write "<request_variable_database>" & RequestDatabase & "</request_variable_database>"
	response.write "<request_variable_ExecBy>" & RequestExecBy & "</request_variable_ExecBy>"
	
	If RequestError = 1 Then
		PostStatus = "Error"
		response.write "<response_server_name></response_server_name>"
		response.write "<response_database_name></response_database_name>"
		response.write "<response_records_inserted></response_records_inserted>"
		response.write "<response_records_updated></response_records_updated>"
	Else
		PostStatus = "Finished"
		set conn=server.CreateObject("ADODB.Connection")
		
		If UCase(Left(RequestServer,6)) = "PHILLY" Then
			ConnString = "driver=SQL Server;server=208.88.49.18,8143;uid=z2t_Subscription_Admin;pwd=get2subscriptions;database=z2t_Maintenance"
			conn.Open ConnString
			SQL = "ha_Products_repl_Table_refresh('" & RequestServer & "', '" & RequestDatabase & "')"
		Else
			conn.Open ReadConnection(RequestServer, RequestDatabase)
			If ID = "" Then
				SQL = "ha_Products_repl_Table_refresh"
			Else
				SQL = "ha_Products_repl_Table_refresh(" & ID & ")"
			End If
		End If
		
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
	
    response.write "</ha_Products_update>"

	'Post to Task Log
	Task = RequestExecBy & " --> ha_Products_update: " & RequestServer & " " & RequestDatabase 
	If ID > "" Then
		Task = Task & " ID=" & ID
	End If
	Call TaskLogWrite("", TaskName, Task, PostStatus)
	
%>

