<!DOCTYPE html>

<html>
<head>
  <% PageHeading = "Post Time Off Request" %>
  <title><%=PageHeading %></title>
  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/Config.asp"-->
  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
</head>
<body>
  <%
    Function fixaps (string)
      fixaps = replace(string, "'", "''")
    End Function
    Function iif(condition,result,alternative)
      If condition Then iif=result Else iif=alternative
    End Function

    Dim strSQL, result, rs: Set rs = Server.CreateObject("ADODB.Recordset")

    If Request("iMode") = "Add" Then
      strSQL = "SELECT * FROM ha_Inventory"
      rs.Open strSQL, conncasper10, 1, 3 
      rs.AddNew
    Else
      
      strSQL = "SELECT * FROM ha_Inventory WHERE ID = " & Request("ID")
      rs.Open strSQL, connCasper10, 1, 3
    End If
      
      If Not Request("MachineName") = "" Then rs("MachineName") = Request("MachineName")
      rs("Description") = Request("Description")
      rs("Vendor") = Request("Vendor")
      rs("Warranty") = Request("Warranty")
      rs("DatePurchased") = Request("DatePurchased")
      rs("ModelNumber") = Request("ModelNumber")
      rs("SerialNumber") = Request("SerialNumber")
      rs("AssignedTo") = Request("AssignedTo")
      rs("Location") = Request("Location")
      rs("Condition") = Request("Condition")
      If Not Request("Notes") = "" Then rs("Notes") = fixaps(Request("Notes"))
      rs.Update
      rs.Close

  %>
  <script type="text/javascript">
    alert('Record has been <%=Request("iMode")%>ed.');
    window.opener.location.href = window.opener.location.href;
    window.close()
  </script>
</body>
</html>
<%
  Function iif(condition, consequent, alternative)
    If condition Then iif = consequent Else iif = alternative
  End Function
%>