<!--#include virtual="ha_backoffice/Servers/TableDistribution/includes/TableDistribution_TaskLog_write.inc"-->
<!--#include virtual="ha_backoffice/Servers/TableDistribution/includes/TableDistribution_Functions.inc"-->

<%
	Dim rs
    Dim SQL
	Dim ConnString
    Dim conn
	Dim RequestServer
	Dim RequestDatabase
	Dim RequestExecBy
	Dim RequestError
	Dim PostStatus
		
	RequestError = 0
	ReadRequestVariables
	
    response.ContentType = "text/xml"
    response.write "<?xml version='1.0' encoding='ISO-8859-1'?>"
    response.write "<ha_Products_status>"
	response.write "<request_variable_server>" & RequestServer & "</request_variable_server>"
	response.write "<request_variable_database>" & RequestDatabase & "</request_variable_database>"
	response.write "<request_variable_ExecBy>" & RequestExecBy & "</request_variable_ExecBy>"

	If RequestError = 1 Then
		PostStatus = "Error"
		response.write "<response_server_name></response_server_name>"
		response.write "<response_database_name></response_database_name>"
		response.write "<response_record_count></response_record_count>"
	Else
		PostStatus = "Finished"
		set conn=server.CreateObject("ADODB.Connection")

		If UCase(Left(RequestServer,6)) = "PHILLY" Then
			ConnString = "driver=SQL Server;server=208.88.49.18,8143;uid=z2t_Subscription_Admin;pwd=get2subscriptions;database=z2t_Maintenance"
			conn.Open ConnString
			SQL = "ha_Products_status('" & RequestServer & "', '" & RequestDatabase & "')"
		Else
			ConnString = "driver=SQL Server;server=127.0.0.1,7043;uid=davewj2o;pwd=get2it;database=ha_Processes"
			conn.Open ConnString
			SQL = "TableDistribution_ha_Products_status('" & RequestServer & "', '" & RequestDatabase & "')"
			'conn.Open ReadConnection(RequestServer, RequestDatabase)
			'SQL = "ha_Products_status"
		End If

		set rs=server.createObject("ADODB.Recordset")
		response.write "<request_SQL>" & SQL & "</request_SQL>"
		rs.open SQL, conn, 3, 3, 4

		If not rs.eof then
			response.write "<response_server_name>" & rs("ServerName") & "</response_server_name>"
			response.write "<response_database_name>" & rs("DatabaseName") & "</response_database_name>"
			response.write "<response_record_count>" & rs("RecordCount") & "</response_record_count>"
		End If

		rs.Close
		set rs = nothing
	End If
	
	response.write "</ha_Products_status>"

	'Post to Task Log
	Task = ExecBy & " --> ha_Products_Status: " & RequestServer & " " & RequestDatabase
	Call TaskLogWrite("", "", Task, "")
	
%>

