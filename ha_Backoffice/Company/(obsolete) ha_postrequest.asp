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

    If Request("rMode") = "Add" Then
      'strSQL = "SELECT MAX(ID) AS MaxID FROM ha_EmpTimeRequest"
      'rs.Open strSQL, objcon, 2, 3
      'Session("ID") = rs("MaxID")
      'rs.Close
      'strSQL = "INSERT INTO ha_EmpTimeRequest (RequestFrom,RequestTo,Reason,Duration,DepartureTime,RequestedBy,RequestedOn) VALUES ('" & Request("RequestFrom") & "','"      If Request("RequestTo") = "" Then strSQL = strSQL & Request("RequestFrom") Else strSQL & Request("RequestTo")
      'strSQL = strSQL &  "','" & fixaps(Request("Reason") & "','" & Request("Duration") & "','"
      'If Request("Duration") = "Partial" Then strSQL = strSQL & Request("DepartureTime")
      'strSQL = strSQL & "','" & fixaps(Request("RequestedBy")) & "','" & Now & "')"
      strSQL = "SELECT * FROM ha_EmpTimeRequest"
      rs.Open strSQL, connCasper10, 1, 3 
      rs.AddNew
      'rs("ID") = Session("ID")
      rs("RequestedBy") = fixaps(Request("RequestedBy"))
      rs("RequestedOn") = Now
    Else
      'strSQL = "UPDATE ha_EmpTimeRequest SET RequestFrom = '" & Request("RequestFrom") & "', RequestTo = '" & iif(Request("RequestTo") = "", Request("RequestFrom"), Request("RequestTo")"', Reason = '" & fixaps(Request("Reason") & "', Duration = '" & Request("Duration") & "'," & iif(Request("Duration") = "Partial", " DepartureTime '" & Request("DepartureTime") & "' ", NULL) & " WHERE ID = " & Request("ID")
      strSQL = "SELECT * FROM ha_EmpTimeRequest WHERE ID = " & Request("ID")
      rs.Open strSQL, connCasper10, 1, 3
    End If
    'Response.write strSQL & chr(13)
    connCasper10.Execute strSQL

    'If Request("rMode") = "Edit" and Session("ha_login") = "lrowlands" Then
    '  strSQL = "UPDATE ha_EmpTimeRequest SET IsApproved = '" & Request("IsApproved") & "', AppReason = '" & fixaps(Request("AppReason")) & "' WHERE ID = " & Request("ID")
    '  conn.Execute strSQL
    'End If
      rs("RequestFrom") = Request("RequestFrom")
      If Request("RequestTo") = "" Then 
        rs("RequestTo") = Request("RequestFrom") 
      Else 
        rs("RequestTo") = Request("RequestTo")
      End If
      rs("Reason") = fixaps(Request("Reason"))
      rs("Duration") = Request("Duration")
      If Not Request("DepartureTime") = "" Then rs("DepartureTime") = Request("DepartureTime")
      If Request("rMode") = "Edit" and Session("ha_login") = "lrowlands" Then
        rs("IsApproved") = Request("IsApproved")
        rs("AppReason") = fixaps(Request("AppReason"))
      End If
      rs.Update
      rs.Close

    
      Dim emailOriginate, emailBody, emailTarget, emailCopy, emailSubject

      If Request("rmode") = "Add" Then
        emailTarget = "lrowlands@harvestamerican.com"
        emailSubject = "Time Off Request - " & Request("RequestedBy")
        emailOriginate = "mailserver@harvestamerican.com"
        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.Open "SELECT * FROM ha_EmpAccounts WHERE FirstName = '" & Request("RequestedBy") & "'", connCasper10, 1, 3
        emailCopy = rs("Email")
        rs.Close
        emailBody = "Greetings Lucinda, " & chr(13) & "There is a new time off request from " & Request("RequestedBy") & " available for viewing at /ha_BackOffice/Company/ha_timerequest.asp.  Please check it at your soonest convenience!" & chr(13) & chr(13) & "This is a system message.  Please do not reply to this message."
      Else
        If Session("ha_login") = "lrowlands" Then
          rs.Open "SELECT * FROM ha_EmpAccounts WHERE FirstName = '" & Request("RequestedBy") & "'", connCasper10, 1, 3
          emailTarget = rs("FirstName") & " " & fixaps(rs("LastName")) & "<" & rs("Email") & ">"
          emailSubject = "RE: Your Time Off Request"
          emailOriginate = "Lucinda Rowlands <lrowlands@harvestamerican.com>"
          emailCopy = "Lucinda Rowlands <lrowlands@harvestamerican.com>"
          rs.Close
          emailBody = "Greetings " & Request("RequestedBy") & ", " & chr(13)
          
          If Request("IsApproved") = 1 Then
            emailBody = emailBody & "I am pleased to inform you that I have approved your request for time off. Enjoy!" & chr(13) & chr(13) & "Sincerely," & chr(13) & chr(13) & "Lucinda"
          Else
            emailBody = emailBody & "Unfortunately, I am unable to approve your request for time off for the following reason: " & chr(13) & chr(13) & Request("AppReason") & chr(13) & chr(13) & "I hope you understand why I had to do this.  If you do not, please do not hesitate to get in touch with me." & chr(13) & chr(13) & "Sincerely," & chr(13) & chr(13) & "Lucinda"
          End If
        Else
          emailTarget = "lrowlands@harvestamerican.com"
          emailSubject = "Time Off Request - " & Request("RequestedBy")
          emailOriginate = "mailserver@harvestamerican.com"
          rs.Open "SELECT * FROM ha_EmpAccounts WHERE FirstName = '" & Request("RequestedBy") & "'", connCasper10, 1, 3
          emailCopy = rs("Email")
          rs.Close
          emailBody = "Greetings Lucinda, " & chr(13) & "There is a new time off request from " & Request("RequestedBy") & " that was corrected and is available for viewing at /ha_BackOffice/Company/ha_timerequest.asp.  Please check it at your soonest convenience!" & chr(13) & chr(13) & "This is a system message.  Please do not reply to this message."
        End If
      End If

      strSQL = "util_SendMail '" & emailOriginate & "', '" & emailTarget & "', '" & emailCopy & "', '" & fixaps(emailSubject) & "', '" & fixaps(emailBody) & "'"
      'Response.write strSQL & chr(13)
      connCasper10.Execute strSQL
      connCasper10.close
  %>
  <script type="text/javascript">
    alert('Record has been <%=Request("rMode")%>ed.');
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