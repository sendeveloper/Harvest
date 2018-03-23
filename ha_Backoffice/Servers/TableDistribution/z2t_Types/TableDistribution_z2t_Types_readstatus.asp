<!--#include virtual="ha_backoffice/Servers/TableDistribution/includes/TableDistribution_TaskLog_write.inc"-->
<!--#include virtual="ha_backoffice/Servers/TableDistribution/includes/TableDistribution_Functions.inc"-->

<%
	Dim rs
    Dim SQL
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
    response.write "<z2t_Types_status>"
	response.write "<request_variable_server>" & RequestServer & "</request_variable_server>"
	response.write "<request_variable_database>" & RequestDatabase & "</request_variable_database>"
	response.write "<request_variable_ExecBy>" & RequestExecBy & "</request_variable_ExecBy>"
	
	If RequestError = 1 Then
		PostStatus = "Error"
		response.write "<response_server_name></response_server_name>"
		response.write "<response_database_name></response_database_name>"
		response.write "<response_record_count></response_record_count>"
		response.write "<response_class_count></response_class_count>"
	Else
		PostStatus = "Finished"
		set conn=server.CreateObject("ADODB.Connection")
		
		conn.open "FILEDSN=C:\Inetpub\DSN\" & RequestServer & "SQLServerTableDistribution.DSN"
		'conn.Open "DRIVER=ODBC Driver 11 for SQL Server;server=Casper06.HarvestAmerican.net,7643;uid=davewj2o;pwd=get2it;database=ha_TableDistribution"
		'conn.Open "DRIVER=ODBC Driver 11 for SQL Server;server=10.119.55.114,7643;uid=davewj2o;pwd=get2it;database=ha_TableDistribution"
				
		set rs=server.createObject("ADODB.Recordset")
		SQL = "ha_TableDistribution_z2t_Types_status('" & RequestDatabase & "')"		
		response.write "<request_SQL>" & SQL & "</request_SQL>"
		rs.open SQL, conn
		
		If not rs.eof then
			response.write "<response_server_name>" & rs("ServerName") & "</response_server_name>"
			response.write "<response_database_name>" & rs("DatabaseName") & "</response_database_name>"
			response.write "<response_record_count>" & rs("RecordCount") & "</response_record_count>"
			response.write "<response_class_count>" & rs("ClassCount") & "</response_class_count>"
		End If

		rs.Close
		set rs = nothing
	End If
	
	response.write "</z2t_Types_status>"
	
	'Post to Task Log
	Task = ExecBy & " --> z2t_Types_Status: " & RequestServer & " " & RequestDatabase
	Call TaskLogWrite("", "", Task, "")

%>

