<!--#include virtual="includes/connection.asp"-->
<!--#include virtual="ha_backoffice/includes/FunctionsPermissions.inc"-->

<html>
<head>
    <title>Post</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />

</head>
<body>

<%
    Dim rs, strSQL

    Set rs = Server.CreateObject("ADODB.Recordset")
    	
    strSQL = "SELECT * FROM ha_Employees WHERE ID = " & Request("ID")
    rs.Open strSQL, connCasper10, 2, 3
	rs("EditedBy") = Session("ha_shortname")
    rs("EditedDate") = Now

	If HasPermission("Edit", 7) = True Then
		rs("UserName")		= ltrim(rtrim(Request("StaffUN")))
	End If
	If HasPermission("Edit", 8) = True Then
		rs("Password")		= ltrim(rtrim(Request("StaffPWD")))
	End If
	  
    rs.Update
    rs.Close
%>

<script type="text/javascript">
    window.opener.location.href = window.opener.location.href;
    window.close()
</script>

</body>	
</html>
