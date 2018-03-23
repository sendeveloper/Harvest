<!--#include virtual="includes/connection.asp"-->

<html>
<head>
    <title>Post</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />

</head>
<body>

<%
    Dim rs, strSQL

    Set rs = Server.CreateObject("ADODB.Recordset")
    	
    If Request("ID") = "0" Then
      strSQL = "SELECT * FROM ha_Domains"
      rs.Open strSQL, connCasper10, 2, 3
      rs.AddNew
	  rs("CreatedBy") = Session("ha_shortname")
      rs("CreatedDate") = Now
    Else
      strSQL = "SELECT * FROM ha_Domains WHERE ID = " & Request("ID")
      rs.Open strSQL, connCasper10, 2, 3
	  rs("EditedBy") = Session("ha_shortname")
      rs("EditedDate") = Now
    End If

	  
	rs("SubName") = ltrim(rtrim(Request("SubName")))
	rs("DomainName") = ltrim(rtrim(Request("DomainName")))
	rs("Purpose") = ltrim(rtrim(Request("Purpose")))
	rs("Class") = ltrim(rtrim(Request("Class")))
	rs("ServerName") = ltrim(rtrim(Request("ServerName")))
	rs("SecureSite") = ltrim(rtrim(Request("SecureSite")))
	rs("FTPLogin") = ltrim(rtrim(Request("FTPLogin")))
	rs("FTPPassword") = ltrim(rtrim(Request("FTPPassword")))
	rs("DomainRegister") = ltrim(rtrim(Request("DomainRegister")))
	rs("DB") = ltrim(rtrim(Request("DB")))
	rs("Notes") = ltrim(rtrim(Request("Notes")))

    rs.Update
    rs.Close
%>

<script type="text/javascript">
    window.opener.location.href = window.opener.location.href;
    window.close()
</script>

</body>	
</html>
