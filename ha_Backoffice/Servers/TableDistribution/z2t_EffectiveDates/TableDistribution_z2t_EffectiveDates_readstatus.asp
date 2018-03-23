<!--#include virtual="ha_backoffice/Servers/TableDistribution/includes/TableDistribution_TaskLog_write.inc"-->
<!--#include virtual="ha_backoffice/Servers/TableDistribution/includes/TableDistribution_Functions.inc"-->

<%
	Dim rs
    Dim SQL
    Dim conn
	Dim connDallas
	Dim RequestServer
	Dim RequestDatabase
	Dim RequestExecBy
	Dim RequestError
	Dim PostStatus
		
	RequestError = 0
	ReadRequestVariables
	
    response.ContentType = "text/xml"
    response.write "<?xml version='1.0' encoding='ISO-8859-1'?>"
    response.write "<z2t_EffectiveDates_status>"
	response.write "<request_variable_server>" & RequestServer & "</request_variable_server>"
	response.write "<request_variable_database>" & RequestDatabase & "</request_variable_database>"
	response.write "<request_variable_ExecBy>" & RequestExecBy & "</request_variable_ExecBy>"
	
	If RequestError = 1 Then
		PostStatus = "Error"
		response.write "<response_server_name></response_server_name>"
		response.write "<response_database_name></response_database_name>"
		response.write "<response_record_count></response_record_count>"
		response.write "<response_maximum_AccountID></response_maximum_AccountID>"
	Else
		PostStatus = "Finished"
		set conn=server.CreateObject("ADODB.Connection")

		conn.open "FILEDSN=C:\Inetpub\DSN\" & RequestServer & "SQLServerTableDistribution.DSN"
				
		set rs=server.createObject("ADODB.Recordset")
		SQL = "ha_TableDistribution_z2t_EffectiveDates_status('" & RequestDatabase & "')"		
		response.write "<request_SQL>" & SQL & "</request_SQL>"
		rs.open SQL, conn
		
		
		'ServerFarm = UCase(Left(RequestServer,6))
		'If ServerFarm = "FRANK0" Then
		'	conn.open "FILEDSN=C:\Inetpub\DSN\" & RequestServer & "SQLServerTableDistribution.DSN"
				
		'	set rs=server.createObject("ADODB.Recordset")
		'	SQL = "ha_TableDistribution_z2t_EffectiveDates_status('" & RequestDatabase & "')"		
		'	response.write "<request_SQL>" & SQL & "</request_SQL>"
		'	rs.open SQL, conn
		'Else
		'	Select Case ServerFarm
		'	Case "CASPER"
		'		ConnString = "driver=SQL Server;server=10.119.55.118,7043;uid=z2t_Subscription_Admin;pwd=get2subscriptions;database=ha_PublishedTables"
		'		conn.Open ConnString
		'		SQL = "z2t_EffectiveDates_status('" & RequestServer & "', '" & RequestDatabase & "')"
		'	Case "PHILLY"
		'		ConnString = "driver=SQL Server;server=208.88.49.18,8143;uid=z2t_Subscription_Admin;pwd=get2subscriptions;database=z2t_Maintenance"
		'		conn.Open ConnString
		'		SQL = "z2t_EffectiveDates_status('" & RequestServer & "', '" & RequestDatabase & "')"
		'	Case Else
		'		conn.Open ReadConnection(RequestServer, RequestDatabase)
		'		SQL = "z2t_EffectiveDates_status"
		'	End Select

		'	set rs=server.createObject("ADODB.Recordset")
		'	response.write "<request_SQL>" & SQL & "</request_SQL>"
		'	rs.open SQL, conn, 3, 3, 4
		'End If

		If not rs.eof then
			response.write "<response_server_name>" & rs("ServerName") & "</response_server_name>"
			response.write "<response_database_name>" & rs("DatabaseName") & "</response_database_name>"
			response.write "<response_record_count>" & rs("RecordCount") & "</response_record_count>"
		End If

		rs.Close
		set rs = nothing
	End If
	
	response.write "</z2t_EffectiveDates_status>"

	'Post to Task Log
	Task = ExecBy & " --> z2t_EffectiveDates_Status: " & RequestServer & " " & RequestDatabase
	Call TaskLogWrite("", "", Task, "")
	
%>

