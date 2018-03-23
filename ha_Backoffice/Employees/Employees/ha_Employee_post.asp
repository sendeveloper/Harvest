<!--#include virtual="includes/connection.asp"-->
<!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->
<!--#include virtual="ha_backoffice/includes/FunctionsPermissions.inc"-->

<html>
<head>
    <title>Employee Edit Post</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />

</head>
<body>

<%
    Dim rs, strSQL

    Set rs = Server.CreateObject("ADODB.Recordset")
	
    If Request("ID") = "0" Then
      strSQL = "SELECT * FROM ha_Employees"
      rs.Open strSQL, connCasper10, 2, 3
      rs.AddNew
	  rs("CreatedBy") = Session("ha_shortname")
      rs("CreatedDate") = Now
    Else
      strSQL = "SELECT * FROM ha_Employees WHERE ID = " & Request("ID")
      rs.Open strSQL, connCasper10, 2, 3
	  rs("EditedBy") = Session("ha_shortname")
      rs("EditedDate") = Now
    End If
	
		rs("FirstName")		= ltrim(rtrim(Request("FirstName")))
		rs("LastName")		= ltrim(rtrim(Request("LastName")))
		rs("Email")			= ltrim(rtrim(Request("Email")))
		If HasPermission("Edit", 9) = True Then
			DOB = ltrim(rtrim(Request("DOB")))
			rs("DOB") 		= iif(DOB = "", null, DOB)
		End If
		rs("AddressLine1")	= ltrim(rtrim(Request("AddressLine1")))
		rs("AddressLine2")	= ltrim(rtrim(Request("AddressLine2")))
		rs("City")			= ltrim(rtrim(Request("City")))
		rs("State")			= ltrim(rtrim(Request("State")))
		rs("PostalCode")	= ltrim(rtrim(Request("PostalCode")))
		rs("WorkLocation")	= ltrim(rtrim(Request("WorkLocation")))
		rs("Phone")			= ltrim(rtrim(Request("HomePhone")))
		rs("Mobile")		= ltrim(rtrim(Request("MobilePhone")))
		rs("Skype")			= ltrim(rtrim(Request("Skype")))
		rs("HarvestID") 	= iif(Request("HarvestID") = "", null, Request("HarvestID"))
		
	  
    rs.Update
    rs.Close
	
	strSQL = "UPDATE ha_Employees SET EmpID = ID WHERE isnull(EmpID,0) = 0"
	connCasper10.Execute strSQL
	
%>

<script type="text/javascript">
    window.opener.location.href = window.opener.location.href;
    window.close()
</script>

</body>	
</html>
