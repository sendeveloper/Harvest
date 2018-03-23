  <!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->
  
<html>
<head>
	<title>Harvest American Backoffice - Products Post</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />

</head>
<body>
<%
    Dim rs, strSQL

  	Set connLocal = Server.CreateObject("ADODB.Connection")
	connLocal.Open "driver=SQL Server; server=127.0.0.1,7043; uid=davewj2o; pwd=get2it; database=ha_backoffice"

    Set rs = server.createObject("ADODB.Recordset")
		
    If Request("ID") = "0" or Request("ID") = "" Then
      strSQL = "SELECT * FROM ha_Server_Maintenance_Log"
      rs.Open strSQL, connLocal, 2, 3
      rs.AddNew
	  rs("CreatedBy") = Session("ha_shortname")
      rs("CreatedDate") = Now
    Else
      strSQL = "SELECT * FROM ha_Server_Maintenance_Log WHERE ID = " & Request("ID")
      rs.Open strSQL, connLocal, 2, 3
	  rs("EditedBy") = Session("ha_shortname")
      rs("EditedDate") = Now
    End If
	
		rs("ServerName") 				= Request("ServerName")
		rs("Task") 						= Request("Task")
		rs("Type") 						= Trim(Request("Type"))
		rs("Engineer")					= Request("Engineer")
		rs("Notes") 					= Trim(Request("Notes"))
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
