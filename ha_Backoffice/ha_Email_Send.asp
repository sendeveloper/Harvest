<!--#include virtual="/ha_backoffice/includes/config.asp"-->
<!--#include virtual="/includes/connection.asp"-->

<html>
<head>
    <title>Send E-mail</title>
</head>

<%
    Dim SQL

    'Send E-mail
    SQL = "util_SendMail '" & Request("SendFrom") & "', " & _
        "'" & Request("SendTo") & "', " & _
	"'" & Request("Copy") & "', " & _
	"'" & Request("Subject") & "', " & _
        "'" & fixAps(Request("noteTextArea")) & "'"
''    response.Write(sql)
	
	connCasper10.Execute SQL

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
'response.write sql
    ObjCon.Execute SQL


    FUNCTION fixAps( ByVal theString )
        fixAps = Replace( theString, "'", "''" )
    END FUNCTION
%>

<script>
    window.opener.location.href = window.opener.location.href;
    window.close()
</script>

</html>

