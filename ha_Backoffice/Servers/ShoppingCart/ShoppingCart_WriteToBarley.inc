
<%
	Dim connBarley

	set connBarley=server.CreateObject("ADODB.Connection")
	 
 	'connBarley.Open "driver=SQL Server;server=68.178.202.54;uid=davewj2o;pwd=get2it;database=ha_prod" 'Barley1
 	connBarley.Open "driver=SQL Server;server=66.119.50.228,7843;uid=davewj2o;pwd=get2it;database=ha_prod" 'Casper08

	SQL = "INSERT ni_UploadFromShoppingCart (" & qfields & ", CreatedBy, CreatedDate) " & _
		"VALUES (" & qReq & ", 'ShoppingCart_WriteToBarley.inc', GETDATE())"
    connBarley.execute SQL
	
	
    'rs.Open "SELECT * FROM ni_UploadFromShoppingCart",connBarley,1,3
    'rs.AddNew

	'	rs("m_code") = Request("m_code")
'		rs("inv_num") = Request("inv_num")
'		rs("order_num") = Request("order_num")
'		rs("date") = Request("date")
'		rs("time") = Request("time")
'		rs("subtotal") = Request("subtotal")
'		rs("ship_charge") = Request("ship_charge")
'		rs("sales_tax") = Request("sales_tax")
'		rs("cod_charge") = Request("cod_charge")
'		rs("min_charge") = Request("min_charge")
'		rs("vol_discount") = Request("vol_discount")
'		rs("pro_discount") = Request("pro_discount")
'		rs("pro_code") = Request("pro_code")
'		rs("total") = Request("total")
'		rs("pay_method") = Request("pay_method")
'		rs("cc_type") = Request("cc_type")
'		rs("cc_num") = Request("cc_num")
'		rs("cc_number") = Request("cc_number")
'		rs("cc_month") = Request("cc_month")
'		rs("cc_year") = Request("cc_year")
'		rs("gateway_approval") = Request("gateway_approval")
'		rs("gateway_avs") = Request("gateway_avs")
'		rs("gateway_trans") = Request("gateway_trans")
'		rs("gateway_cvv2") = Request("gateway_cvv2")
'		rs("b_fname") = Request("b_fname")
'		rs("b_lname") = Request("b_lname")
'		rs("b_addr1") = Request("b_addr1")
'		rs("b_addr2") = Request("b_addr2")
'		rs("b_city") = Request("b_city")
'		rs("b_state") = Request("b_state")
'		rs("b_zip") = Request("b_zip")
'		rs("b_country") = Request("b_country")
'		rs("b_phone") = Request("b_phone")
'		rs("b_email") = Request("b_email")
'		rs("b_company") = Request("b_company")
'		rs("b_fax") = Request("b_fax")
'		rs("b_web") = Request("b_web")
'		rs("comments") = Request("comments")
'		rs("ship_method") = Request("ship_method")
'		rs("order_weight") = Request("order_weight")
'		rs("s_company") = Request("s_company")
'		rs("s_fname") = Request("s_fname")
'		rs("s_lname") = Request("s_lname")
'		rs("s_addr1") = Request("s_addr1")
'		rs("s_addr2") = Request("s_addr2")
'		rs("s_city") = Request("s_city")
'		rs("s_state") = Request("s_state")
'		rs("s_zip") = Request("s_zip")
'		rs("s_country") = Request("s_country")
'		rs("s_phone") = Request("s_phone")
'		rs("s_fax") = Request("s_fax")
'		rs("s_web") = Request("s_web")
'		rs("s_email") = Request("s_email")
'		rs("s_addr_type") = Request("s_addr_type")
'		rs("cust_track") = Request("cust_track")
'		rs("cust_ip") = Request("cust_ip")
'		rs("item1") = Request("item1")
'		rs("item2") = Request("item2")
'		rs("item3") = Request("item3")
'		rs("item4") = Request("item4")
'		rs("item5") = Request("item5")
'		rs("item6") = Request("item6")
'		rs("item7") = Request("item7")
'		rs("item8") = Request("item8")
'		rs("CreatedBy") = "ShoppingCart_WriteToBarley.inc"
'		rs("CreatedDate") = now()
'    rs.Update
'    rs.Close

	connBarley.Close
	
    response.write "Successful posting to Barley.ha_prod.ni_UploadFromShoppingCart" & vbCrLf
	
	'Post to Task Log
	Call TaskLogWrite("Casper10", "UploadShoppingCart.asp", "ShoppingCart_WriteToBarley.inc", "Finished")	

%>