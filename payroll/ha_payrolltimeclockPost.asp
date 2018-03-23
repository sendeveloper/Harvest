<!--#include virtual="includes/connection.asp"-->

<html>
<head>
    <title>Employee Hours Post</title>

    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />

</head>
<body>
<%

			Dim rs
			Dim SQL
			
			if Request("PayDate") <> "" then
               Session("ePayDate") = Request("PayDate")
            end if
						
			set rs = server.createObject("ADODB.Recordset")
			SQL="ha_PayrollUpdate('" & Session("ePayDate") &"')"

			rs.open SQL,connCasper10,2,3	
			
		    Response.Redirect "http://www.harvestamerican.info/ha_Backoffice/Company/ha_Payroll.asp"

%>

	
<p>&nbsp;</p>
</body>
</html>