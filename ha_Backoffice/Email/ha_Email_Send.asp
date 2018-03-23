
<%
response.Buffer=true
%>
<!--#include virtual="ha_backoffice/includes/config.asp"-->
<!--#include virtual="ha_backoffice/includes/connection.asp"-->

<html>
<head>
    <title>Send E-mail</title>


<%

    Dim SQL
'        "'" & Request("SendTo") & "', " & _
    'Send E-mail
    SQL = "util_SendMail '" & Request("SendFrom") & "', " & _
	      "'nazishabbasi@gmail.com', " & _
		  	"'" & Request("Copy") & "', " & _
	"'" & Request("Subject") & "', " & _
        "'" & fixAps(Request("noteTextArea")) & "'"
''    response.Write(sql)
	
	connEmail.Execute SQL

    'Report to Activity
    SQL = "INSERT INTO ni_Activity " & _
        "(UserName, ActivityType, Data1, Data2, ForAccountID, SessionID, CreatedByIP, CreatedBy, CreatedDate) " & _
        "VALUES " & _
            "('" & Session("Login") & "', " & _
            "'Email Sent', " & _
            "'" & Request("Subject") & "', " & _
            "'Raffle Ticket', " & _
            "'" & Session("AccountID") & "', " & _
            "'', " & _
            "'', " & _
            "'ha_Email_Send.asp', " & _
            "GETDATE())"
''response.write sql
   ObjCon.Execute SQL


    FUNCTION fixAps( ByVal theString )
        fixAps = Replace( theString, "'", "''" )
    END FUNCTION
%>

<script>
// window.opener.location.href = window.opener.location.href;
    window.close()
</script>
</head>
</html>

