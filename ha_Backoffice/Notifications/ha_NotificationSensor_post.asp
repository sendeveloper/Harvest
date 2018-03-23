  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/Config.asp"-->
  
<!DOCTYPE html>

<html>
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
  <title>Harvest American Backoffice - Post ToDo</title>
</head>

<body>

<%
    'response.write Request("AssignedTo") & "<br>"
    'response.write Request("ManagedBy")
	'response.end
	
	Dim rs, strSQL
		
	If isnull(Request("Priority")) or ltrim(Request("Priority")) = "" or Request("Priority") = "0" Then
		Priority = null
	Else
		Priority = Request("Priority")
	End If
	
    ' open recordset
    Set rs = Server.CreateObject("ADODB.Recordset")
    
    If Request("ID") = "0" Then
      strSQL = "SELECT * FROM ha_NotificationSensors"
      rs.Open strSQL, connCasper10, 2, 3
      rs.AddNew
	  rs("CreatedBy") = Session("ha_shortname")
      rs("CreatedDate") = Now
    Else
      strSQL = "SELECT * FROM ha_NotificationSensors WHERE ID = " & Request("ID")
      rs.Open strSQL, connCasper10, 2, 3
	  rs("EditedBy") = Session("ha_shortname")
      rs("EditedDate") = Now
    End If

		rs("Title") = ltrim(rtrim(Request("Title")))
		rs("Description") = ltrim(rtrim(Request("Description")))
		rs("Increment") = Request("Increment")
		rs("ServerName") = Request("ServerName")
		rs("Category") = Request("Category")
		rs("Status") = Request("Status")
    rs.Update
    rs.Close
%>
  
  <script type="text/javascript">
    window.opener.location.href = window.opener.location.href;
    window.close();
  </script>
  
</body>
</html>
