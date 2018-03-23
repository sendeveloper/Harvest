<!--#include virtual="includes/connection.asp"-->

<html>
<head>
    <title>Inventory Edit Post</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />

</head>
<body>

<%
    Dim rs, strSQL

    Set rs = Server.CreateObject("ADODB.Recordset")
	
    If Request("ID") = "0" Then
      strSQL = "SELECT * FROM ha_Inventory"
      rs.Open strSQL, connCasper10, 2, 3
      rs.AddNew
    Else
      strSQL = "SELECT * FROM ha_Inventory WHERE ID = " & Request("ID")
      rs.Open strSQL, connCasper10, 2, 3
    End If
	
		rs("Type")	 	    = ltrim(rtrim(Request("Type")))
		rs("Description")	= ltrim(rtrim(Request("Description")))
		rs("Condition")		= ltrim(rtrim(Request("Condition")))
		rs("AssignedTo")	= ltrim(rtrim(Request("AssignedTo")))
		rs("Quantity")		= ltrim(rtrim(Request("Quantity")))
		
	  
    rs.Update
    rs.Close
	
	'strSQL = "UPDATE ha_Inventory SET ID = ID WHERE isnull(ID,0) = 0"
	connCasper10.Execute strSQL
	
%>

<script type="text/javascript">
    window.opener.location.href = window.opener.location.href;
    window.close()
</script>

</body>	
</html>
