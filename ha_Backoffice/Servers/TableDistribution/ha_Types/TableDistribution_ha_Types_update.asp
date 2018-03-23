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
		TaskName = "TableDistribution_ha_Types_update (Full)"
	Else
		ID = Request("ID")
		TaskName = "TableDistribution_ha_Types_update (" & ID & ")"
	End If
	
    response.ContentType = "text/xml"
    response.write "<?xml version='1.0' encoding='ISO-8859-1'?>"
    response.write "<ha_Types_update>"
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
		
		If UCase(Left(RequestServer,6)) = "BARLEY" Then
			conn.Open ReadConnection(RequestServer, RequestDatabase)
			SQL = "ha_Types_repl_Table_refresh"
		Else
			ConnString = "driver=SQL Server;server=127.0.0.1,7043;uid=davewj2o;pwd=get2it;database=ha_Processes"  'Casper10 (Local)
			conn.Open ConnString
			If ID = "" Then			
				SQL = "TableDistribution_ha_Types_update('" & RequestServer & "', '" & RequestDatabase & "')"
			Else
				SQL = "TableDistribution_ha_Types_update('" & RequestServer & "', '" & RequestDatabase & "', " & ID & ")"
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
	
    response.write "</ha_Types_update>"

	'Post to Task Log
	Task = RequestExecBy & " --> ha_Types_update: " & RequestServer & " " & RequestDatabase & " " & AccountID
	Call TaskLogWrite("", TaskName, Task, PostStatus)
	
%>

