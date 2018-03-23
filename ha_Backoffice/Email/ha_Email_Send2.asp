<!--#include virtual="/includes/config.asp"-->
<!--#include virtual="/includes/connection.asp"-->
<%
    Dim SQL
	
    set rs = server.createObject("ADODB.Recordset")
	
	    SQL = Request("eSQL")

    rs.open SQL, connEmail'', 3, 3, 4

    eSendTo = rs("Email")
    eSendFrom = rs("SendFrom")
    eSubject = rs("Subject")
    eText = rs("eText")

    'Send E-mail
    SQL = "util_SendMail '" & rs("SendFrom") & "', " & _
        "'" & Request("eSendTo")  & "', " & _
	"'" & Request("eCopy") & "', " & _
	"'" & rs("Subject") & "', " & _
        "'" & fixAps(rs("eText")) & "'"
    connEmail.Execute SQL
''response.Write(sql)
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
    objCon.Execute SQL

response.write "Email Sent successfully."

    FUNCTION fixAps( ByVal theString )
        fixAps = Replace( theString, "'", "''" )
    END FUNCTION
%>

