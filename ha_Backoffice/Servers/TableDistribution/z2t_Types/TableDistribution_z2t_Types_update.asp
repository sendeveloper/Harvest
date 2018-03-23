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


    response.ContentType = "text/xml"
    response.write "<?xml version='1.0' encoding='ISO-8859-1'?>"
    response.write "<z2t_Types_update>"
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
		
		conn.open "FILEDSN=C:\Inetpub\DSN\" & RequestServer & "SQLServerTableDistribution.DSN"
		'conn.Open "DRIVER=ODBC Driver 11 for SQL Server;server=10.119.55.114,7643;uid=davewj2o;pwd=get2it;database=ha_TableDistribution"
		
		set rs=server.createObject("ADODB.Recordset")
		SQL = "ha_TableDistribution_z2t_Types_update('" & RequestDatabase & "')"		
		'SQL = "ha_TableDistribution_z2t_Types_status('" & RequestDatabase & "')"		
		response.write "<request_SQL>" & SQL & "</request_SQL>"
		rs.open SQL, conn
					
		If not rs.eof then
			response.write "<response_server_name>" & rs("ServerName") & "</response_server_name>"
			response.write "<response_database_name>" & rs("DatabaseName") & "</response_database_name>"
			response.write "<response_records_inserted>" & rs("RecordsInserted") & "</response_records_inserted>"
			response.write "<response_records_updated>" & rs("RecordsUpdated") & "</response_records_updated>"
		End If
		
		rs.Close
		set rs = nothing
	End If
	
    response.write "</z2t_Types_update>"

	'Post to Task Log
	Task = ExecBy & " --> z2t_Types_update: " & RequestServer & " " & RequestDatabase & " " & AccountID
	Call TaskLogWrite("", "", Task, PostStatus)

%>

