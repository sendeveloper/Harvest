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
    
    If Request("AccomplishmentDate") = "" or IsNull(Request("AccomplishmentDate")) Then
		eDate = null
	Else
		eDate = Request("AccomplishmentDate")
    End If
	
    If Request("ID") = "0" Then
      strSQL = "SELECT * FROM ha_Accomplishments"
      rs.Open strSQL, connCasper10, 2, 3
      rs.AddNew
	  rs("CreatedBy") = Session("ha_shortname")
      rs("CreatedDate") = Now
    Else
      strSQL = "SELECT * FROM ha_Accomplishments WHERE ID = " & Request("ID")
      rs.Open strSQL, connCasper10, 2, 3
	  rs("EditedBy") = Session("ha_shortname")
      rs("EditedDate") = Now
    End If

	  rs("Date") = eDate
	  rs("Description") = ltrim(rtrim(Request("Description")))
    rs.Update
    rs.Close
%>

<script type="text/javascript">
    window.opener.location.href = window.opener.location.href;
    window.close()
</script>

</body>	
</html>
