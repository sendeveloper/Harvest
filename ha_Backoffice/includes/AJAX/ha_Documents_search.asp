<%
	Dim rs
    Dim SQL
	Dim connCasper10

	If isnull(Request("f")) Then
		f = ""
	Else
		f = Request("f")
	End If
	
	If f > "" Then
	
		set connCasper10=server.CreateObject("ADODB.Connection")
		connCasper10.Open "driver=SQL Server;server=66.119.55.118,7043;uid=davewj2o;pwd=get2it;database=ha_BackOffice"
		
		SQL = "ha_Document_search('" & f & "')"
		
		set rs=server.createObject("ADODB.Recordset")
		rs.open SQL, connCasper10, 3, 3, 4

		s = "<table cellspacing='2' cellpadding='2'" & vbCrLf
		
		If rs.EOF Then
            s = s & "<tr>"
            s = s & "    <td colspan='5'>"
            s = s & "        <CENTER>No file with that search name</CENTER>"
            s = s & "    </td>"
            s = s & "</tr>"
		Else
			LineCount = 0
			do while not rs.eof
				s = s & "<tr><td class='fileRow'>" & _
					"<a href='javascript:clickAttachSelect(" & _
					cStr(rs("ImageID")) & _
					");' class='fileRow'>" & _
					rs("FileName") & _
					"</a></td></tr>" & vbCrLf
				rs.movenext
			loop
			
			s = s & "</table>" & vbCrLf
			Response.write s
		End If
		
		rs.Close
		set rs = nothing
		
	End If
	
%>

