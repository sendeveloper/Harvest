<%
	Dim connCasper10: Set connCasper10=server.CreateObject("ADODB.Connection")
    Dim strSQL
	connCasper10.Open "driver=SQL Server;server=66.119.55.118,7043;uid=davewj2o;pwd=get2it;database=ha_Backoffice"

	If Request("delTable") = "ha_Documents" Then
		strSQL = "ha_Document_delete(" & Request("delID") & ", '" & Request("delUser") & "')"
	Else
		strSQL = "UPDATE " & Request("delTable") & " SET DeletedDate = GETDATE(), DeletedBy = '" & Request("delUser") & "' WHERE ID = " & Request("delID")
	End If

	'response.write strSQL
	connCasper10.Execute strSQL
%>

