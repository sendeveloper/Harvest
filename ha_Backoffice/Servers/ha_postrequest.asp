<!DOCTYPE html>

<html>
<head>
  <% PageHeading = "Post Time Off Request" %>
  <title><%=PageHeading %></title>
  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/Config.asp"-->
  <!--#include file="sql.asp"-->
  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
</head>
<body>
  <%
    Function iif(condition,result,alternative)
      If condition Then iif=result Else iif=alternative
    End Function

    Dim strSQL, result

    If Request("rMode") = "Add" Then
      strSQL = "INSERT INTO [ha_BackOffice].[dbo].[ha_EmpTimeRequest] ([RequestFrom], " & iif(Not Request("RequestTo") = "", "[RequestTo], ", "") & "[Reason], [RequestedBy], [RequestedOn]" & iif(Not Request("IsApproved") = "", ", [IsApproved], [AppReason]", "") & ") VALUES ('" & Request("RequestFrom") & "', " & iif(Not Request("RequestTo") = "", "'" & Request("RequestTo") & "', ","") & "'" & Request("Reason") & "', '" & Request("RequestedBy") & "', '" & Date & "'"& iif(Request("IsApproved") = "", ")", ", " & Request("IsApproved") & ",'" & Request("AppReason") & "')")
    Else
      strSQL = "UPDATE ha_EmpTimeRequest SET RequestFrom = '" & Request("RequestFrom") & "', " & iif(Not Request("RequestTo") = "", "RequestTo = '" & Request("RequestTo") & "', ", "") & "Reason = '" & Request("Reason") & "'" & iif(Not Request("IsApproved") = "", ", IsApproved = '" & Request("IsApproved") & "', AppReason = '" & Request("AppReason") & "'", "") & " WHERE ID = " & Request("ID")
    End If
    'Response.write strSQL
    objcon.Execute strSQL

    
      Dim emailOriginate, emailBody, emailTarget, emailCopy, emailSubject, rs

      If Request("rmode") = "Add" Then
        emailTarget = "lrowlands@harvestamerican.com"
        emailSubject = "Time Off Request - " & Request("RequestedBy")
        emailOriginate = "mailserver@harvestamerican.com"
        Set rs = Sql("SELECT * FROM ha_EmpAccounts WHERE FirstName = '" & Request("RequestedBy") & "'")
        emailCopy = rs("Email")
        rs.Close
        emailBody = "Greetings Lucinda, " & chr(13) & "There is a new time off request from " & Request("RequestedBy") & " available for viewing at /ha_BackOffice/Company/ha_timerequest.asp.  Please check it at your soonest convenience!" & chr(13) & chr(13) & "This is a system message.  Please do not reply to this message."
      Else
        Set rs = Sql("SELECT * FROM ha_EmpAccounts WHERE FirstName = '" & Request("RequestedBy") & "'")
        emailTarget = rs("FirstName") & " " & rs("LastName") & "<" & rs("Email") & ">"
        emailSubject = "RE: Your Time Off Request"
        emailOriginate = "Lucinda Rowlands <lrowlands@harvestamerican.com>"
        emailCopy = "Lucinda Rowlands <lrowlands@harvestamerican.com>"
        rs.Close
        emailBody = "Greetings " & Request("RequestedBy") & ", " & chr(13) & iif(Request("IsApproved") = 1,"I am pleased to inform you that I have approved your request for time off. Enjoy!" & chr(13) & chr(13) & "Sincerely," & chr(13) & chr(13) & "Lucinda","Unfortunately, I am unable to approve your request for time off for the following reason: " & chr(13) & chr(13) & Request("AppReason") & chr(13) & chr(13) & "I hope you understand why I had to do this.  If you do not, please do not hesitate to get in touch with me." & chr(13) & chr(13) & "Sincerely," & chr(13) & chr(13) & "Lucinda")
      End If

      strSQL = "util_SendMail '" & emailOriginate & "', '" & emailTarget & "', '" & emailCopy & "', '" & emailSubject & "', '" & emailBody & "'"
      objcon.Execute strSQL
  %>
  <script type="text/javascript">
    alert('Record has been <%=Request("rMode")%>ed.');
    window.opener.location.href = window.opener.location.href;
    window.close()
  </script>
</body>
</html>
