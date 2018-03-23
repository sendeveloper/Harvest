<%

    '// URL to which to send XML data
	URL = "/ha_backoffice/servers/ShoppingCart/woocommerce/uploadwoocommerce.pl"

    Response.Write "Sending XML data to " & URL & "<br/>"

    information = "<Send><UserName>Colt</UserName><PassWord>Taylor</PassWord><Data>100</Data></Send>"
	
    Set xmlhttp = server.Createobject("MSXML2.ServerXMLHTTP")
    xmlhttp.Open "POST", url, false
    xmlhttp.setRequestHeader "Content-Type", "text/xml" 
    xmlhttp.send information

    '// Report the response from the called page
	Response.Write "<br/><br/>"
    response.write "Response received:<hr/><span style='color:blue'>" & xmlhttp.ResponseText & "</span><hr/>"

%>
