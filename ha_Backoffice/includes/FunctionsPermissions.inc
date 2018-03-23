<%
	Bullets8 = "&bull;&bull;&bull;&bull;&bull;&bull;&bull;&bull;"
	
	Function HasPermission(pType, pID)
	
		'Requires a Casper10 Connection
		'pType = [read, edit, or delete]  A permission of ReadEdit means either/or
		'pID = [PermissionID]
		
		Set rsHP = Server.CreateObject("ADODB.Recordset")
		
        rsHP.open "ha_Employee_Permission_read('" & pType & "', " & Session("ha_EmployeeID") & ", " & pid & ")", connCasper10, 0, 1
		
		If rsHP("result") = 1 Then HasPermission = True Else HasPermission = False
		
		rsHP.close
		set rsHP = Nothing
		
	End Function
	
	Function CheckLoadPermitted (pID)
	
		Set rsCLP = Server.CreateObject("ADODB.Recordset")
		
        rsCLP.open "ha_Employee_Permission_read('read', " & Session("ha_EmployeeID") & ", " & pid & ")", connCasper10, 0, 1
		
		If rsCLP("result") <> 1 Then 
			Response.Redirect "/ha_BackOffice/ha_Home.asp"
		End If
		
		rsCLP.close
		set rsCLP = Nothing
		
	End Function
%>