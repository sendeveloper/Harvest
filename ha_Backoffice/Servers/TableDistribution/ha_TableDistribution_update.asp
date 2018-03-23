<!--#include virtual="ha_backoffice/Servers/TableDistribution/includes/TableDistribution_TaskLog_write.inc"-->
<!--#include virtual="ha_backoffice/Servers/TableDistribution/includes/TableDistribution_Functions.inc"-->

<%
    Dim rs
    Dim SQL
    Dim conn
	Dim ExecBy
	Dim RequestServer
	Dim RequestDatabase
	Dim RequestTable
	Dim RequestClass
	Dim RequestExecBy
	Dim RequestError
	Dim PostStatus

	RequestError = 0
	ReadRequestVariables	
	
	z2tID = "0"
	HarvestID = "0"
	
	If isnull(Request("z2tID")) or Request("z2tID") = "" Then
		If isnull(Request("HarvestID")) or Request("HarvestID") = "" Then
			'Full distribution
			TaskName = "TableDistribution_z2t_Accounts_Subscriptions_Update (Full)"
		Else
			'Account distribution
			HarvestID = Request("HarvestID")
			TaskName = "TableDistribution_z2t_Accounts_Subscriptions_Update (Account)"
		End If
	Else
		'Subscription account distribution
		z2tID = Request("z2tID")
		TaskName = "TableDistribution_z2t_Accounts_Subscriptions_Update (Subscription)"
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

    response.ContentType = "text/xml"
    response.write "<?xml version='1.0' encoding='ISO-8859-1'?>"
    response.write "<ha_TableDistribution_update>"
	response.write "<request_variable_server>" & RequestServer & "</request_variable_server>"
	response.write "<request_variable_database>" & RequestDatabase & "</request_variable_database>"
	response.write "<request_variable_table>" & RequestTable & "</request_variable_table>"
	response.write "<request_variable_objects_class>" & RequestClass & "</request_variable_objects_class>"
	response.write "<request_variable_ExecBy>" & RequestExecBy & "</request_variable_ExecBy>"
	response.write "<request_variable_Start>" & RequestStart & "</request_variable_Start>"
	response.write "<request_variable_End>" & RequestEnd & "</request_variable_End>"
	
	If RequestError = 1 Then
		PostStatus = "Error"
		response.write "<response_server_name></response_server_name>"
		response.write "<response_database_name></response_database_name>"
		response.write "<response_records_inserted></response_records_inserted>"
		response.write "<response_records_updated></response_records_updated>"
		response.write "<response_records_deleted></response_records_deleted>"
	Else
		PostStatus = "Finished"
		set conn=server.CreateObject("ADODB.Connection")
		set rs=server.createObject("ADODB.Recordset")
		
		conn.open "FILEDSN=C:\Inetpub\DSN\" & RequestServer & "SQLServerTableDistribution.DSN"
		SQL = "ha_TableDistribution_Table_update('" & RequestDatabase & "', " & _
			"'" & RequestTable & "', '" & RequestClass & "', " & _
			RequestStart & ", " & RequestEnd & ")"
				
		response.write "<request_SQL>" & SQL & "</request_SQL>"
		rs.open SQL, conn', 3, 3, 4

		If not rs.eof then
			response.write "<response_server_name>" & rs("ServerName") & "</response_server_name>"
			response.write "<response_database_name>" & rs("DatabaseName") & "</response_database_name>"
			response.write "<response_table_name>" & rs("TableName") & "</response_table_name>"
			response.write "<response_class_name>" & rs("ClassName") & "</response_class_name>"
			response.write "<response_data_source>" & rs("DataSource") & "</response_data_source>"
			response.write "<response_records_inserted>" & rs("RecordsInserted") & "</response_records_inserted>"
			response.write "<response_records_updated>" & rs("RecordsUpdated") & "</response_records_updated>"
			response.write "<response_records_deleted>" & rs("RecordsDeleted") & "</response_records_deleted>"
			response.write "<response_record_start>" & rs("RecordStart") & "</response_record_start>"
			response.write "<response_record_end>" & rs("RecordEnd") & "</response_record_end>"
		End If
		
		rs.Close
		set rs = nothing
	End If
	
    response.write "</ha_TableDistribution_update>"
	
	'Post to Task Log
	Task = RequestExecBy & " --> ha_TableDistribution_update: " & RequestServer & " " & RequestDatabase
	If z2tID <> "0" Then
		Task = Task & " z2tID=" & z2tID
	End If
	If HarvestID <> "0" Then
		Task = Task & " HarvestID=" & HarvestID
	End If
	Call TaskLogWrite("", TaskName, Task, PostStatus)
	
%>

