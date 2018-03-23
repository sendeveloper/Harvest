<%

Dim xmlhttp
Dim DataToSend
Dim strUser
strUser = "sample"
Dim strPassword
strPassword = "password"
Dim strZip
strZip = "90210"

Dim postUrl
postUrl = "https://api.zip2tax.com/TaxRate-USA.xml?username="&strUser&"&password="&strPassword&"&zip="&strZip

Set xmlhttp = server.Createobject("MSXML2.ServerXMLHTTP")
xmlhttp.Open "GET",postUrl,false
xmlhttp.send
Response.ContentType = "charset=utf-8"

Set xmlDoc = CreateObject("MSXML2.DOMDocument") 'creating the parser object
xmlDoc.LoadXml(xmlhttp.responseText)

Set itemList = xmlDoc.SelectNodes("/z2tLookup/addressInfo/addresses/address")
For Each itemAttrib In itemList
    
    Dim zip_code : zip_code = itemAttrib.SelectSingleNode("zipCode").text
    Dim place : place = itemAttrib.SelectSingleNode("place").text
    Dim county : county = itemAttrib.SelectSingleNode("county").text
    Dim state : state = itemAttrib.SelectSingleNode("state").text
Next

Set itemTaxList = xmlDoc.SelectNodes("/z2tLookup/addressInfo/addresses/address/salesTax/rateInfo")
For Each itemAttrib In itemTaxList
    
    Dim taxRate : taxRate = itemAttrib.SelectSingleNode("taxRate").text
Next

Dim count:  count = 0
Set itemTaxList = xmlDoc.SelectNodes("/z2tLookup/addressInfo/addresses/address/notes/noteDetail")
For Each itemAttrib In itemTaxList  
    if count = 1 then
        Dim shipping_note : shipping_note = itemAttrib.SelectSingleNode("note").text
    end if
    count = count + 1
Next

Dim shipping : shipping = 0
if shipping_note = "Shipping charges are not taxable" then
        shipping = 1
end if

Response.Write "Zip Code: " & zip_code &  vbCrLf
Response.Write "Sales Tax Rate: " & taxRate &  vbCrLf
Response.Write "Post Office City: " & place &  vbCrLf
Response.Write "County: " & county &  vbCrLf
Response.Write "State: " & state &  vbCrLf
Response.Write "Shipping Taxable: " & shipping &  vbCrLf


set xmlhttp = nothing
%>
