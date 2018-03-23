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

	If isnull(Request("z2tID")) or Request("z2tID") = "" Then
		z2tID = "0"
		TaskName = "TableDistribution_z2t_Accounts_Subscriptions_addons_Update (Full)"
	Else
		z2tID = Request("z2tID")
		TaskName = "TableDistribution_z2t_Accounts_Subscriptions_addons_Update (Partial)"
	End If
	

    response.ContentType = "text/xml"
    response.write "<?xml version='1.0' encoding='ISO-8859-1'?>"
    response.write "<z2t_Accounts_Subscriptions_addons_update>"
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
		Server.ScriptTimeout = 240
		conn.CommandTimeout = 240

		conn.open "FILEDSN=C:\Inetpub\DSN\" & RequestServer & "SQLServerTableDistribution.DSN"
		SQL = "ha_TableDistribution_z2t_Accounts_Subscriptions_addons_update('" & RequestDatabase & "', " & _
			z2tID & ", " & RequestStart & ", " & RequestEnd & ")"
		
		'If UCase(Left(RequestServer,6)) = "PHILLY" Then
		'	connString = "driver=SQL Server;server=208.88.49.18,8143;uid=z2t_Subscription_Admin;pwd=get2subscriptions;database=z2t_Maintenance"
		'Else
		'	connString = "driver=SQL Server;server=127.0.0.1,7043;uid=davewj2o;pwd=get2it;database=ha_Processes"
		'End If
		'conn.Open ConnString

		set rs=server.createObject("ADODB.Recordset")
	
		'SQL = "TableDistribution_z2t_Accounts_Subscriptions_addons_update('" & RequestServer & "', '" & RequestDatabase & "', " & z2tID & ")"		
		response.write "<request_SQL>" & SQL & "</request_SQL>"
		rs.open SQL, conn,,,4
		
		If not rs.eof then
			response.write "<response_server_name>" & rs("ServerName") & "</response_server_name>"
			response.write "<response_database_name>" & rs("DatabaseName") & "</response_database_name>"
			response.write "<response_records_inserted>" & rs("RecordsInserted") & "</response_records_inserted>"
			response.write "<response_records_updated>" & rs("RecordsUpdated") & "</response_records_updated>"
		End If
		
		rs.Close
		set rs = nothing
	End If
	
    response.write "</z2t_Accounts_Subscriptions_addons_update>"

	'Post to Task Log
	Task = ExecBy & " --> z2t_Accounts_Subscriptions_addons_Update: " & RequestServer & " " & RequestDatabase
	If z2tID <> "0" Then
		Task = Task & " z2tID=" & z2tID
	End If
	Call TaskLogWrite("", TaskName, Task, PostStatus)
%>

