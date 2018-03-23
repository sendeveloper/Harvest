<!--#include virtual="ha_backoffice/Servers/TableDistribution/includes/TableDistribution_TaskLog_write.inc"-->
<!--#include virtual="ha_backoffice/Servers/TableDistribution/includes/TableDistribution_Functions.inc"-->

<%
	Dim rs
    Dim SQL
	Dim ConnString
    Dim conn
	Dim connDallas
	Dim RequestServer
	Dim RequestDatabase
	Dim RequestExecBy
	Dim RequestError
	Dim PostStatus
		
	RequestError = 0
	ReadRequestVariables
		
	If isnull(Request("t")) or Request("t") = ""  Then
		RequestError = 1
	Else
		RequestTable = Request("t")
	End If
	
    response.ContentType = "text/xml"
    response.write "<?xml version='1.0' encoding='ISO-8859-1'?>"
    response.write "<z2t_Activity_status>"
	response.write "<request_variable_server>" & RequestServer & "</request_variable_server>"
	response.write "<request_variable_database>" & RequestDatabase & "</request_variable_database>"
	response.write "<request_variable_ExecBy>" & RequestExecBy & "</request_variable_ExecBy>"
	
	If RequestError = 1 Then
		PostStatus = "Error"
		response.write "<response_server_name>|</response_server_name>"
		response.write "<response_database_name>|</response_database_name>"
		response.write "<response_record_count>|</response_record_count>"
		response.write "<response_location_1>|</response_location_1>"
		response.write "<response_location_2>|</response_location_2>"
	Else
		PostStatus = "Finished"
		set conn=server.CreateObject("ADODB.Connection")

		conn.open "FILEDSN=C:\Inetpub\DSN\" & RequestServer & "SQLServerTableDistribution.DSN"
				
		set rs=server.createObject("ADODB.Recordset")
		SQL = "ha_TableDistribution_z2t_Activity_status('" & RequestDatabase & "', '" & RequestTable & "')"		
		response.write "<request_SQL>" & SQL & "</request_SQL>"
		rs.open SQL, conn

		'AltServer = "Philly01"
		'conn.Open ReadConnection(RequestServer, RequestDatabase)
			
		'If UCase(Left(RequestServer,6)) = "PHILLY" THEN
		'	SQL = "z2t_Activity_status('" & RequestServer & "', '" & RequestDatabase & "', '" & RequestTable & "')"
		'Else	
		'	SQL = "z2t_Activity_status"
		'End If

		'set rs=server.createObject("ADODB.Recordset")
		'response.write "<request_SQL>" & SQL & "</request_SQL>"		
		'rs.open SQL, conn, 3, 3, 4

		If not rs.eof then
			response.write "<response_server_name>" & rs("ServerName") & "</response_server_name>"
			response.write "<response_database_name>" & rs("DatabaseName") & "</response_database_name>"
			response.write "<response_table_name>" & rs("TableName") & "</response_table_name>"
			response.write "<response_SQL>" & SQL & "</response_SQL>"
			response.write "<response_record_count>" & rs("RecordCount") & "</response_record_count>"
			response.write "<response_location_1>" & rs("Location1") & "</response_location_1>"
			response.write "<response_location_2>" & rs("Location2") & "</response_location_2>"
		End If

		rs.Close
		set rs = nothing
	End If
	
	response.write "</z2t_Activity_status>"
	
	
	'Post to Task Log
	Task = ExecBy & " --> z2t Activity Status: " & RequestServer & " " & RequestDatabase
	Call TaskLogWrite("", "", Task, PostStatus)
	
%>

