
<%
	Dim connCasper08

	set connCasper08=server.CreateObject("ADODB.Connection")
	 
 	connCasper08.Open "driver=SQL Server;server=66.119.50.228,7843;uid=davewj2o;pwd=get2it;database=ha_Accounting" 'Casper08

	
	SQL = "INSERT ha_UploadFromShoppingCart (" & qfields & ", CreatedBy, CreatedDate) " & _
		"VALUES (" & qReq & ", 'ShoppingCart_WriteToCasper08.inc', GETDATE())"
    connCasper08.execute SQL

	'Insert the entire query string into a table
	'Query = replace(SQL,"'", "")
	'Response.write Query & vbCrLf
    'SQL = "INSERT ha_UploadFromShoppingCart_querystring (Query, CreatedDate) Values ('" & left(Query,2000) & "', GETDATE())"
    'connCasper08.execute SQL
	
	connCasper08.Close

    response.write "Successful posting to Casper08.ha_Accounting.ha_UploadFromShoppingCart" & vbCrLf
	
	'Post to Task Log
	Call TaskLogWrite("Casper10", "UploadShoppingCart.asp", "ShoppingCart_WriteToCasper08.inc", "Finished")	

%>