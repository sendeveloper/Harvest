<%
	
	Dim LocationCount
	Dim LocationServer(50)
	Dim LocationDatabase(50)
	Dim LocationTable(50)
	Dim LocationFarmSource(50)

	
	Sub ReadObjects(oClass)
	
		Dim connLocal
		connLocal = "Provider=SQLOLEDB;Data Source=127.0.0.1,7043;Initial Catalog=ha_Backoffice;UID=davewj2o;PWD=get2it"
		
		Dim rs
		Dim SQL
		set rs = server.createObject("ADODB.Recordset")

		SQL="ha_Table_Distribution_read('" & oClass & "')"
		rs.open SQL, connLocal, 3, 3, 4
		
		LocationCount = -1
		If Not rs.eof Then
			While Not rs.eof
				LocationCount = LocationCount + 1
				LocationServer(LocationCount) = rs("ServerName")
				LocationDatabase(LocationCount) = rs("DatabaseName")
				LocationTable(LocationCount) = rs("TableName")		
				LocationFarmSource(LocationCount) = rs("FarmSource")		
				rs.MoveNext
			Wend
		End If
		
		rs.Close
	
	End Sub
	
%>

