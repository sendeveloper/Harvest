<%
	FUNCTION TaskLogWrite(CreatedBy, TaskName, Task, Status)
	
		DIM IP, p
		
		'Open Connection
		SET connTaskLog=server.CreateObject("ADODB.Connection")
		connTaskLog.Open "driver=SQL Server;server=127.0.0.1,7043;uid=davewj2o;pwd=get2it;database=ha_Backoffice"
		
		If CreatedBy = "" Then
			IP = Request.ServerVariables("LOCAL_ADDR")
		End If

		If TaskName = "" Then
			TaskName = Request.ServerVariables("PATH_INFO")
			p = InStrRev(TaskName, ".")
			If p > 0 Then
				TaskName =  Left(TaskName, p - 1)
			End If
			p = InStrRev(TaskName, "/")
			If p > 0 Then
				TaskName =  Right(TaskName, Len(TaskName) - p)
			End If
		End If
		
		If Status = "" Then
			Status = "Finished"
		End If
		
		SQL = "ha_TaskLog_write('" & IP & "', '" & CreatedBy & "', '" & TaskName & "', '" & Task & "', '" &  Status & "')"
			 
		connTaskLog.Execute SQL
		
		connTaskLog.Close
		SET connTaskLog = Nothing
	
	END FUNCTION
	
%>

