<%
'strServer = "db2.Zip2Tax.com,8143"
strServer = "db2.Zip2Tax.com"
strDBUsername = "z2t_link"
strDBPassword = "H^2p6~r"

'strDBUsername = "davewj2o"
'strDBPassword = "get2it"
strDatabase = "zip2tax"


'Open the connection
Set conn=server.CreateObject("ADODB.Connection")
conn.Open "driver=SQL Server;server=" & strServer & ";uid=" & strDBUsername & ";pwd=" & strDBPassword & ";database=" & strDatabase

'Assign values to the input variables
strZipCode = "90210"  :  'sample zip code must be between 90001 and 90999
strUsername = "sample"
strPassword = "password"

'Open the recordset using the stored procedure
Set rs = server.CreateObject("ADODB.Recordset")
rs.open "z2t_lookup('" & strZipCode & "', '" & strUsername & "', '" & strPassword & "')", conn, 3, 3, 4

'Read the results
If not rs.EOF then
    Response.write "Zip Code: " & rs("Zip_Code") & "<br />"
    Response.write "Sales Tax Rate: " & rs("Sales_Tax_Rate") & "<br />"
    Response.write "Post Office City: " & rs("Post_Office_City") & "<br />"
    Response.write "County: " & rs("County") & "<br />"
    Response.write "State: " & rs("State") & "<br />"
    Response.write "Shipping Taxable: " & rs("Shipping_Taxable") & "<br />"
End If

'Close the Database
rs.Close
conn.Close

%>