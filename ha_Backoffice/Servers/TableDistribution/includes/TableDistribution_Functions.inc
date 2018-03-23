<%
	FUNCTION ReadConnection(ByVal Server, ByVal Database)	
		If UCase(AltServer) = "PHILLY01" And Left(UCase(Server), 6) = "PHILLY" Then
			Server = AltServer
			Database = "z2t_Maintenance"
		End If
	
		SELECT Case UCase(Server)
		'Case "BARLEY"
		'	ReadConnection = "driver=SQL Server;server=68.178.202.54;uid=davewj2o;pwd=get2it;database=" & Database
		Case "BARLEY2"
			ReadConnection = "driver=SQL Server;server=208.109.189.101;uid=davewj2o;pwd=get2it;database=" & Database
		Case "CAMDEN01"
			ReadConnection = "driver=SQL Server;server=24.97.109.162,6143;uid=davewj2o;pwd=get2it;database=" & Database
		Case "CAMDEN02"
			ReadConnection = "driver=SQL Server;server=24.97.109.163,6243;uid=davewj2o;pwd=get2it;database=" & Database
		Case "CAMDEN03"
			ReadConnection = "driver=SQL Server;server=24.97.109.164,6343;uid=davewj2o;pwd=get2it;database=" & Database
		Case "CAMDEN04"
			ReadConnection = "driver=SQL Server;server=24.97.109.165,6443;uid=davewj2o;pwd=get2it;database=" & Database
		Case "CAMDEN05"
			ReadConnection = "driver=SQL Server;server=24.97.109.166,6543;uid=davewj2o;pwd=get2it;database=" & Database
		Case "CASPER06"
			ReadConnection = "driver=SQL Server;server=66.119.50.226,7643;uid=davewj2o;pwd=get2it;database=" & Database
		Case "CASPER07"
			ReadConnection = "driver=SQL Server;server=66.119.50.227,7743;uid=davewj2o;pwd=get2it;database=" & Database
		Case "CASPER08"
			ReadConnection = "driver=SQL Server;server=66.119.50.228,7843;uid=davewj2o;pwd=get2it;database=" & Database
		Case "CASPER09"
			ReadConnection = "driver=SQL Server;server=66.119.50.229,7943;uid=davewj2o;pwd=get2it;database=" & Database
		Case "CASPER10"
			ReadConnection = "driver=SQL Server;server=66.119.50.230,7043;uid=davewj2o;pwd=get2it;database=" & Database
		Case "PHILLY01"
			ReadConnection = "driver=SQL Server;server=208.88.49.18,8143;uid=z2t_Subscription_Admin;pwd=get2subscriptions;database=" & Database
		Case "PHILLY02"
			ReadConnection = "driver=SQL Server;server=208.88.49.19,8243;uid=z2t_Subscription_Admin;pwd=get2subscriptions;database=" & Database
		Case "PHILLY03"
			ReadConnection = "driver=SQL Server;server=208.88.49.20,8343;uid=z2t_Subscription_Admin;pwd=get2subscriptions;database=" & Database
		Case "PHILLY04"
			ReadConnection = "driver=SQL Server;server=208.88.49.21,8443;uid=z2t_Subscription_Admin;pwd=get2subscriptions;database=" & Database
		Case "PHILLY05"
			ReadConnection = "driver=SQL Server;server=208.88.49.22,8543;uid=z2t_Subscription_Admin;pwd=get2subscriptions;database=" & Database
		Case "FRANK02"
			ReadConnection = "driver=SQL Server;server=212.224.124.196,9243;uid=davewj2o;pwd=get2it;database=" & Database
		Case "FRANK03"
			ReadConnection = "driver=SQL Server;server=212.224.124.197,9343;uid=davewj2o;pwd=get2it;database=" & Database
		End Select
	END FUNCTION
	
	FUNCTION ReadRequestVariables
		If isnull(Request("ExecBy")) or Request("ExecBy") = "" Then
			RequestExecBy = ""
		Else
			RequestExecBy = Request("ExecBy")
		End If
		
		If isnull(Request("s")) or Request("s") = "" Then
			RequestError = 1
		Else
			RequestServer = Request("s")
		End If
		
		If isnull(Request("d")) or Request("d") = "" Then
			RequestError = 1
		Else
			RequestDatabase = Request("d")
		End If
		
		If isnull(Request("t")) or Request("t") = "" Then
		Else
			RequestTable = Request("t")
		End If
		
		If isnull(Request("o")) or Request("o") = "" Then
		Else
			RequestClass = Request("o")
		End If

	END FUNCTION
%>

