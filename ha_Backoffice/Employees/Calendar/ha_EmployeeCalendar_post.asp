<!--#include virtual="includes/connection.asp"-->
<!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->

<html>
<head>
    <title>Employee Calendar Post</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />

</head>
<body>

<%
    Dim rs, strSQL

    Set rs = Server.CreateObject("ADODB.Recordset")
    	
    If Request("ID") = "0" Then
      strSQL = "SELECT * FROM ha_EmpCalendar"
      rs.Open strSQL, connCasper10, 2, 3
      rs.AddNew
	  rs("CreatedBy") = Session("ha_shortname")
      rs("CreatedDate") = Now
    Else
      strSQL = "SELECT * FROM ha_EmpCalendar WHERE ID = " & Request("ID")
      rs.Open strSQL, connCasper10, 2, 3
	  rs("EditedBy") = Session("ha_shortname")
      rs("EditedDate") = Now
    End If

		rs("Type")			= ltrim(rtrim(Request("Type")))
		rs("RequestFrom") 	= iif(Request("RequestFrom") = "", null, Request("RequestFrom"))
		rs("RequestTo") 	= iif(Request("RequestTo") = "", null, Request("RequestTo"))
		rs("Reason")		= ltrim(rtrim(Request("Reason")))
		rs("RequestFor")	= ltrim(rtrim(Request("RequestFor")))
		rs("Duration")		= iif(Request("Duration") = "All Day", "All Day", "Partial Day")
		rs("DepartureTime")	= ltrim(rtrim(Request("DepartureTime")))
		rs("IsApproved")	= iif(Request("Approved") = "Yes", True, False)
	  
    rs.Update
    rs.Close
%>

<script type="text/javascript">
    window.opener.location.href = window.opener.location.href;
    window.close()
</script>

</body>	
</html>
