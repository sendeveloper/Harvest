<!--#include file="Currency2Words.asp"-->

<%
	Dim connCasper10

	set connCasper10=server.CreateObject("ADODB.Connection")
	 
 	connCasper10.Open "driver=SQL Server;server=127.0.0.1,7043;uid=davewj2o;pwd=get2it;database=ha_Backoffice" 'Casper10


    'Send notification to cell phone
    if not isnull(Request("inv_num")) and Request("inv_num") <> "unpaid" and Request("inv_num") > "" then

        bText = "FROM: " & Request("b_city") & ", " & Request("b_state") & vbCRLF & vbCRLF & _
            "TIME: " & left(Request("time") & "     ", 5) & vbCRLF & _
            "AMOUNT: $" & Request("total") & vbCRLF
			
        bTextDave = "FROM: " & Request("b_city") & ", " & Request("b_state") & vbCRLF & vbCRLF & _
            "TIME: " & left(Request("time") & "     ", 5) & vbCRLF & _
            "AMOUNT: " & ConvertCurrencyToEnglish(Request("total")) & vbCRLF

        'Verizon Sample
        'SQL = "util_SendMail('Service@HarvestAmerican.com', '3152645925@vtext.com', '' , 'ni_order', '" & bText & "')"

		'AT&T Sample
        'SQL = "util_SendMail('Service@HarvestAmerican.com', '3153343530@txt.att.net', '', 'z2t_order', '" & bText & "')"
		
		InvNum = Request("inv_num")
        If isnumeric(InvNum) Then
			'Call SendNotification(InvNum, "3157270599@vtext.com", bText) 'Briant
			Call SendNotification(InvNum, "3157097145@vtext.com", bText) 'Dave
			'Call SendNotification(InvNum, "3158061476@txt.att.net", bText) 'Teresa Stevens
			Call SendNotification(InvNum, "3158328623@vtext.com", bText) 'Teresa Farnsworth
		End If
    End If
	
	
	connCasper10.Close

    response.write "Successful sending notification E-mails through Casper10" & vbCrLf
	
	'Post to Task Log
	Call TaskLogWrite("Casper10", "UploadShoppingCart.asp", "ShoppingCart_SendTextNotifications.inc", "Finished")	
	
	
	Sub SendNotification(InvNum, adr, bText)
        if cInt(InvNum) < 10000 then
			'Zip2Tax Orders
            SQL = "util_SendMail('Service@HarvestAmerican.com', '" & adr & "', '', 'Zip Order', '" & bText & "')"
        else
			'Number-it Orders
            SQL = "util_SendMail('Service@HarvestAmerican.com', '" & adr & "', '', 'Harvest Order', '" & bText & "')"
        end if
        connCasper10.Execute(SQL)		
	End Sub
%>