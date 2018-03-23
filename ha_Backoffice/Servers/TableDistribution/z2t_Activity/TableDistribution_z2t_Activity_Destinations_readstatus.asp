<!--#include virtual="ha_backoffice/Servers/TableDistribution/includes/TableDistribution_TaskLog_write.inc"-->
<!--#include virtual="ha_backoffice/Servers/TableDistribution/includes/TableDistribution_Functions.inc"-->

<%
	Dim rs
    Dim SQL
	Dim ConnString
    Dim conn
	Dim connDallas
		
	response.ContentType = "text/xml"
    response.write "<?xml version='1.0' encoding='ISO-8859-1'?>"
    response.write "<activity_destinations>"

	set conn=server.CreateObject("ADODB.Connection")
	set rs=server.createObject("ADODB.Recordset")
	

	
	
	'Philly05
	ConnString = "driver=SQL Server;server=208.88.49.22,8543;uid=davewj2o;pwd=get2it;database=z2t_BackOffice"
	conn.Open ConnString
	
	SQL = "z2t_Activity_counts"
	rs.open SQL, conn, 3, 3, 4

	If not rs.eof then
		response.write "<destination>"
			response.write "<server>Philly05</server>"	
			response.write "<database>z2t_Backoffice</database>"	
			response.write "<count_collection_table>" & rs("Collection") & "</count_collection_table>"
			response.write "<count_annual_table>" & rs("Annual") & "</count_annual_table>"
		response.write "</destination>"
	End If
	
	rs.close
	conn.close
	

	'Casper10
	ConnString = "driver=SQL Server;server=Casper10.HarvestAmerican.net,7043;uid=davewj2o;pwd=get2it;database=z2t_BackOffice"
	conn.Open ConnString
	
	set rs=server.createObject("ADODB.Recordset")
	SQL = "z2t_Activity_counts"
	rs.open SQL, conn, 3, 3, 4

	If not rs.eof then
		response.write "<destination>"
			response.write "<server>Casper10</server>"	
			response.write "<database>z2t_Backoffice</database>"	
			response.write "<count_collection_table>" & rs("Collection") & "</count_collection_table>"
			response.write "<count_annual_table>" & rs("Annual") & "</count_annual_table>"
		response.write "</destination>"
	End If
	
	rs.close
	conn.close

	response.write "</activity_destinations>"
	

	'Post to Task Log
	Task = ExecBy & " --> z2t_Activity_Destinations_Status"
	Call TaskLogWrite("", "", Task, "")

%>

