<%
Response.buffer=true
%>

  <!--#include virtual="includes/connection.asp"-->
  
<html>
<head>
    <title>Document Maintenance Edit Post</title>
</head>

<body>
<%
	
    Dim SQL
 
    SQL = "ha_Document_write(" & _ 
          Request("ID") & ", " & _
          "'" & Request("Type") & "', " & _
          "'" & Request("Business") & "', " & _
          "'" & Request("Department") & "', " & _
          "'" & Session("ha_shortname") & "')" 
    Response.write sql & "<br>"
	connCasper10.Execute SQL
	

%>

<script>
    window.opener.location.href = window.opener.location.href;
    window.close()
</script>
	
</body>
</html>
