<html>
<head>
    <title>Harvest American Backoffice - Requirements Post</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />

</head>
<body>
<%
    Dim rs, strSQL

  	Set connLocal = Server.CreateObject("ADODB.Connection")
	connLocal.Open "driver=SQL Server; server=127.0.0.1,7043; uid=davewj2o; pwd=get2it; database=ha_PublishedTables"

    Set rs = server.createObject("ADODB.Recordset")
	
	strSQL = "SELECT * FROM ha_Products WHERE ID = " & Request("ID")
	rs.Open strSQL, connLocal, 2, 3
	rs("EditedBy") = Session("ha_shortname")
	rs("EditedDate") = Now
	
	rs("RequiresShipping") 				= Request("RequiresShipping")
	rs("RequiresRegistration_rt") 		= Request("RequiresRegistration_rt")
	rs("RequiresRegistration_nm") 		= Request("RequiresRegistration_nm")
	rs("RequiresRegistration_nmpro") 	= Request("RequiresRegistration_nmpro")
	rs("RequiresRegistration_lookup") 	= Request("RequiresRegistration_lookup")
	rs("RequiresEmailTables") 			= Request("RequiresEmailTables")
	rs("RequiresRegistration_initial") 	= Request("RequiresRegistration_initial")
	rs("RequiresRegistration_updates") 	= Request("RequiresRegistration_updates")
	rs("RequiresRegistration_link") 	= Request("RequiresRegistration_link")
	rs("RequiresProcedureSetup") 		= Request("RequiresProcedureSetup")

    rs.Update 
    rs.Close
    connLocal.close
%>

<script type="text/javascript">
    window.opener.location.href = window.opener.location.href;
    window.close()
</script>

</body>	
</html>