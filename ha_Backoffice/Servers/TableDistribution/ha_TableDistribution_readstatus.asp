<!--#include virtual="ha_backoffice/Servers/TableDistribution/includes/TableDistribution_TaskLog_write.inc"-->
<!--#include virtual="ha_backoffice/Servers/TableDistribution/includes/TableDistribution_Functions.inc"-->

<%
	Dim rs
    Dim SQL
	Dim ConnString
    Dim conn
	Dim RequestServer
	Dim RequestDatabase
	Dim RequestTable
	Dim RequestExecBy
	Dim RequestError
	Dim PostStatus
		
	RequestError = 0
	ReadRequestVariables
		
    response.ContentType = "text/xml"
    response.write "<?xml version='1.0' encoding='ISO-8859-1'?>"
    response.write "<ha_TableDistribution_readstatus>"
	response.write "<request_variable_server>" & RequestServer & "</request_variable_server>"
	response.write "<request_variable_database>" & RequestDatabase & "</request_variable_database>"
	response.write "<request_variable_table>" & RequestTable & "</request_variable_table>"
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
		SQL = "ha_TableDistribution_Table_status('" & RequestDatabase & "', '" & RequestTable & "')"		
		response.write "<request_SQL>" & SQL & "</request_SQL>"
		rs.open SQL, conn

		If not rs.eof then
			response.write "<response_server_name>" & rs("ServerName") & "</response_server_name>"
			response.write "<response_database_name>" & rs("DatabaseName") & "</response_database_name>"
			response.write "<response_record_count>" & rs("RecordCount") & "</response_record_count>"
			response.write "<response_minimum_ID>" & rs("MinID") & "</response_minimum_ID>"
			response.write "<response_maximum_ID>" & rs("MaxID") & "</response_maximum_ID>"
		End If

		rs.Close
		set rs = nothing
	End If
	
	response.write "</ha_TableDistribution_readstatus>"
	

	'Post to Task Log
	Task = ExecBy & " --> ha_TableDistribution_readstatus: " & RequestServer & " " & RequestDatabase
	Call TaskLogWrite("", "", Task, "")
	
%>

