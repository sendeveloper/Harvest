  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->

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

		rs("HiredDate")			= iif(Request("HireDate") = "", null, Request("HireDate"))
		rs("TerminationDate")	= iif(Request("TerminationDate") = "", null, Request("TerminationDate"))
		rs("ActiveStatus")		= iif(Request("EmpStatus") = "Active", True, False)
		rs("EmpStatus")			= ltrim(rtrim(Request("EmpType")))
		rs("Contractor")		= iif(Request("Contractor") = "Yes", True, False)
		rs("ProjectSpecialist")	= iif(Request("ProjectSpecialist") = "Yes", True, False)
		rs("PermitTimeClock")	= iif(Request("TimeClock") = "Yes", True, False)
		rs("PermitStaffStatus")	= iif(Request("StaffStatus") = "Yes", True, False)
	  
		rs("EditedBy") = Session("ha_shortname")
		rs("EditedDate") = Now
		
    rs.Update
    rs.Close
%>

<script type="text/javascript">
    window.opener.location.href = window.opener.location.href;
    window.close()
</script>

</body>	
</html>
