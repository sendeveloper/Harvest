<%
	Dim connCasper10: Set connCasper10=server.CreateObject("ADODB.Connection")
    Dim strSQL
	connCasper10.Open "driver=SQL Server;server=66.119.55.118,7043;uid=davewj2o;pwd=get2it;database=ha_Backoffice"

	strSQL = "ha_Employee_Preference_write(" & Session("ha_EmployeeID") & ", '" & Session("ha_shortname") & "', '"  & Request("prefURL") & "', " & _
		"'" & Request("Data1") & "', " & _
		"'" & Request("Data2") & "', " & _
		"'" & Request("Data3") & "', " & _
		"'" & Request("Data4") & "', " & _
		"'" & Request("Data5") & "', " & _
		"'" & Request("Data6") & "', " & _
		"'" & Request("Data7") & "', " & _
		"'" & Request("Data8") & "', " & _
		"'" & Request("Data9") & "')"

	response.write strSQL
	connCasper10.Execute strSQL
%>

