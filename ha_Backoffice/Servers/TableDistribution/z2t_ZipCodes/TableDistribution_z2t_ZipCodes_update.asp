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
	
	TaskName = "TableDistribution_z2t_ZipCodes_update"
		
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

    response.ContentType = "text/xml"
    response.write "<?xml version='1.0' encoding='ISO-8859-1'?>"
    response.write "<z2t_ZipCodes_update>"
	response.write "<request_variable_server>" & RequestServer & "</request_variable_server>"
	response.write "<request_variable_database>" & RequestDatabase & "</request_variable_database>"
	response.write "<request_variable_ExecBy>" & RequestExecBy & "</request_variable_ExecBy>"
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
		set rs=server.createObject("ADODB.Recordset")
		
		conn.open "FILEDSN=C:\Inetpub\DSN\" & RequestServer & "SQLServerTableDistribution.DSN"
		SQL = "ha_TableDistribution_z2t_ZipCodes_update('" & RequestDatabase & "', " & _
				RequestStart & ", " & RequestEnd & ")"
		
		response.write "<request_SQL>" & SQL & "</request_SQL>"
		rs.open SQL, conn

		If not rs.eof then
			response.write "<response_server_name>" & rs("ServerName") & "</response_server_name>"
			response.write "<response_database_name>" & rs("DatabaseName") & "</response_database_name>"
			response.write "<response_records_inserted>" & rs("RecordsInserted") & "</response_records_inserted>"
			response.write "<response_records_updated>" & rs("RecordsUpdated") & "</response_records_updated>"
			response.write "<response_records_deleted>" & rs("RecordsDeleted") & "</response_records_deleted>"
		End If
		
		rs.Close
		set rs = nothing
	End If
	
	response.write "<response_task_name>" & TaskName & "</response_task_name>"
    response.write "</z2t_ZipCodes_update>"


	'Post to Task Log
	Task = RequestExecBy & " --> z2t_SalesTaxRates_Update: " & RequestServer & " " & RequestDatabase
	Task = Task & " Start=" & RequestStart & " End=" & RequestEnd
	Call TaskLogWrite("", TaskName, Task, PostStatus)

%>

