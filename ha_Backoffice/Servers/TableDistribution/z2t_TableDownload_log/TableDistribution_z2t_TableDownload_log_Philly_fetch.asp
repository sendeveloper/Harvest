<!--#include virtual="ha_backoffice/Servers/TableDistribution/includes/TableDistribution_ReadObjects.inc"-->
<!--#include virtual="ha_backoffice/Servers/TableDistribution/includes/TableDistribution_TaskLog_write.inc"-->
<!--#include virtual="ha_backoffice/Servers/TableDistribution/includes/TableDistribution_Functions.inc"-->

<%
	'This asp is use to grab all the activity from various locations
	'on the Philly servers and bring them to one location
	'on Philly01.z2t_Maintenance
	
	Call ReadObjects("z2t_Activity_PhillyFetch")

	Dim rs
    Dim SQL
	Dim ConnString
    Dim conn
	Dim PostStatus
		
	RequestError = 0

	If isnull(Request("ExecBy")) or Request("ExecBy") = "" Then
		ExecBy = ""
	Else
		ExecBy = Request("ExecBy")
	End If
	
    response.ContentType = "text/xml"
    response.write "<?xml version='1.0' encoding='ISO-8859-1'?>"
	response.write "<Philly_activity>"
	
	for i = 1 to LocationCount
		response.write "<Philly_activity_fetch_" & cStr(i) & ">"
		response.write "<TransferDatabaseObject_server>" & LocationServer(i) & "</TransferDatabaseObject_server>"
		response.write "<TransferDatabaseObject_database>" & LocationDatabase(i) & "</TransferDatabaseObject_database>"
		response.write "<TransferDatabaseObject_table>" & LocationTable(i) & "</TransferDatabaseObject_table>"

		PostStatus = "Finished"
		set conn=server.CreateObject("ADODB.Connection")

		ConnString = "driver=SQL Server;server=208.88.49.18,8143;uid=z2t_Subscription_Admin;pwd=get2subscriptions;database=z2t_Maintenance" 'Philly01
		conn.Open ConnString
		
		SQL = "z2t_Activity_fetch('" & LocationServer(i) & "', '" & LocationDatabase(i) & "', '" & LocationTable(i) & "')"
		set rs=server.createObject("ADODB.Recordset")
		'Response.write SQL
		rs.open SQL, conn, 3, 3, 4

		RecordsMoved = 0
		If not rs.eof then
			RecordsMoved = rs("RecordsMoved")
			response.write "<SQL>" & SQL & "</SQL>"
			response.write "<response_server_name>" & rs("ServerName") & "</response_server_name>"
			response.write "<response_database_name>" & rs("DatabaseName") & "</response_database_name>"
			response.write "<response_table_name>" & rs("TableName") & "</response_table_name>"
			response.write "<response_records_moved>" & RecordsMoved & "</response_records_moved>"
		End If

		rs.Close
		set rs = nothing

		'Post to Task Log
		Task = ExecBy & " --> Casper10.TableDistribution_z2t_Activity_Philly_fetch.asp --> Philly01.z2t_Maintenance.z2t_Activity_fetch: " & _
			LocationServer(i) & "." & LocationDatabase(i) & "." & LocationTable(i) & " - " & _
			"Records: " & cStr(RecordsMoved)
		Call TaskLogWrite("", "TableDistribution_z2t_Activity_Philly_fetch", Task, PostStatus)

		response.write "</Philly_activity_fetch_" & cStr(i) & ">"
	Next

	response.write "</Philly_activity>"
	
%>

