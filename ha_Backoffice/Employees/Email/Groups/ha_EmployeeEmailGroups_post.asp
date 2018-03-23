  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/Config.asp"-->
  
<!DOCTYPE html>

<html>


<head>
  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
  <title>Harvest American Backoffice - Post Employee Email Groups</title>
</head>


<body>

<%
    'response.write Request("EmpID") & "<br>"  
	
	
	Dim rs, strSQL
	eEmployeeID = "0"
	Set rs = Server.CreateObject("ADODB.Recordset")
	eEmployeeID = Request("EmployeeID")
	IF Request("EmployeeID") = ""  THEN 
		eEmployeeID = "0"
	END IF
	
	strSQL = "ha_Employee_EmailGroups_Members_save(" & Request("GroupID") & " , '" & eEmployeeID &"' , '" & Session("ha_shortname") & "')"
	
	response.write strSQL & "<br>"
	'response.end
	rs.Open strSQL, connLocal, 3, 3, 4
	
	'response.end
%>
  
  <script type="text/javascript">
    window.opener.location.href = window.opener.location.href;
    window.close();
  </script>
  
</body>
</html>
