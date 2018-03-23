
<%
	Dim connCasper08

	set connCasper08=server.CreateObject("ADODB.Connection")
	 
 	connCasper08.Open "driver=SQL Server;server=66.119.50.228,7843;uid=ShoppingCart_write;pwd=ha_Shop42^b;database=ha_Accounting" 'Casper08

	
	'Insert the entire query string into a table
	

	
    Dim objXmlRequest
    Set objXmlRequest = Server.CreateObject("MSXML2.DOMDOCUMENT.3.0")

    IF objXmlRequest.Load (Request) THEN

      'GET THE REQUEST FROM CLIENT
      strQuery = "//UserName"
      strQuery = "//"
      Set oNode = objXmlRequest.selectSingleNode(strQuery)
      strActionName = oNode.Text
      response.write "success! user name is " & strActionName  & "<br>"
	  'Set Root = objXmlRequest.documentElement 
	  'Response.Write Server.HTMLEncode(objXmlRequest) & "<br><br>"
	  'Response.Write Root.text & "<br><br>"
	  
	set objXMLRoot = objXmlRequest.selectSingleNode(strQuery)

	if not (objXMLRoot is nothing) then
	  'for each objNode in objXMLRoot.childNodes()
	   'Put what is needed in here
	   'response.write objNode.text & "<br>"
	  'next
	end if
	  
	  Query = objXmlRequest.text

    ELSE

        Response.Write "Failed to load XML file, reason: " & objXmlRequest.parseError.reason 
		Query = Request.ServerVariables ("QUERY_STRING")
    END IF

		
	Query = left(Query,2000)
	response.write Query & "<br>"
	
	information = "<Send><UserName>Colt</UserName><PassWord>Taylor</PassWord><Data>100</Data></Send>"
	information = Query
	
    SQL = "INSERT ha_UploadFromShoppingCart_raw (ResponseText, ResponseXML, CreatedDate) Values ('" & Query & "', '" & information & "', GETDATE())"
    connCasper08.execute SQL
	
	connCasper08.Close

    response.write "Successful posting to Casper08.ha_Accounting.ha_UploadFromShoppingCart" & vbCrLf
	
	'Post to Task Log
	Call TaskLogWrite("Casper10", "UploadWoocommerce.asp", "ShoppingCart_Woocommerce_WriteToCasper08.inc", "Finished")	

%>