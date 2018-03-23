<%

	Function UserTrackingAdd(Operation, Data1, Data2)

		Dim TrackURL
		Dim TrackDomain

		'Get TrackURL
		TrackURL = Request.ServerVariables("HTTP_REFERER")
		If len(TrackURL) > 250 Then
			TrackURL = Left(TrackURL, 250)
		End If

		'Get TrackDomain
		p = instr(TrackURL, "//")
		If p > 0 Then
			TrackDomain = Right(TrackURL,Len(TrackURL) - (p + 1))
		Else
			TrackDomain = TrackURL
		End If

		p = instr(TrackDomain, "/")
		If p > 0 Then
			TrackDomain = Left(TrackDomain,p - 1)
		End If
		
		Source = ""
		Country = ""
		IpBigInt = 0

		sqlText = "ha_UserTracking_add('" & Session.SessionID & "', " & _
			"'" & Session("ha_login") & "', " & _
			"'" & Request.ServerVariables("PATH_INFO") & "', " & _
			"'" & Request.ServerVariables("REMOTE_ADDR") & "', " & _
			"'" & TrackURL & "', " & _
			"'" & TrackDomain & "', " & _
			"'" & Source & "', " & _
			IpBigInt & ", " & _
			"'" & Country & "', " & _
			"'" & Operation & "', " & _
			"'" & Data1 & "', " & _
			"'" & Data2 & "')"
		'response.write sqlText & "<br>"
		connLocal.Execute(sqlText) 
		
	End Function
%>