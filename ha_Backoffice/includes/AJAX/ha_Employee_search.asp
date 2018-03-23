<%
	Dim rs
    Dim SQL
	Dim connCasper10

	If isnull(Request("f")) Then
		f = ""
	Else
		f = Request("f")
	End If
	
	'If f > "" Then
	
		set connLocal=server.CreateObject("ADODB.Connection")
		connLocal.Open "driver=SQL Server;server=127.0.0.1,7043;uid=davewj2o;pwd=get2it;database=ha_BackOffice"
		
		SQL = "ha_Employee_search('" & f & "')"
		
		set rs=server.createObject("ADODB.Recordset")
		rs.open SQL, connLocal, 3, 3, 4

		s = "<table cellspacing='2' cellpadding='2'" & vbCrLf
		
		If rs.EOF Then
            s = s & "<tr>"
            s = s & "    <td colspan='5' align='center'>"
            s = s & "        No employee with that search name"
            s = s & "    </td>"
            s = s & "</tr>"
		Else
			Do While Not rs.eof
				s = s & "<tr>"
				s = s & "<td class='fileRow'>" & _
					"<a href='javascript:clickAssigneeSelect(" & _
					cStr(rs("ID")) & _
					");' class='fileRow'>" & _
					rs("ShortName") & _
					"</a></td>"
				s = s & "<td class='fileRow'>" & _
					"<a href='javascript:clickAssigneeSelect(" & _
					cStr(rs("ID")) & _
					");' class='fileRow'>" & _
					rs("FullName") & _
					"</a></td>"					
				s = s & "</tr>" & vbCrLf
				rs.movenext
			Loop			
		End If
		
		s = s & "</table>" & vbCrLf
		Response.write s		
		
		rs.Close
		set rs = nothing
		
	'End If
	
%>

