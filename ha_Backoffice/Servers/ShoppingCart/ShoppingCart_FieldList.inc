
<%
	Dim f(100)
	
	
	f(1) 	= "m_code"
	f(2) 	= "inv_num"
	f(3) 	= "order_num"
	f(4) 	= "date"
	f(5) 	= "time"
	f(6) 	= "subtotal"
	f(7) 	= "ship_charge"
	f(8) 	= "sales_tax"
	f(9) 	= "cod_charge"
	f(10) 	= "min_charge"
	f(11) 	= "vol_discount"
	f(12) 	= "pro_discount"
	f(13) 	= "pro_code"
	f(14) 	= "total"
	f(15) 	= "pay_method"
	f(16) 	= "cc_type"
	f(17) 	= "cc_num"
	f(18) 	= "cc_number"
	f(19) 	= "cc_month"
	f(20) 	= "cc_year"
	f(21) 	= "gateway_approval"
	f(22) 	= "gateway_avs"
	f(23) 	= "gateway_trans"
	f(24) 	= "gateway_cvv2"
	f(25) 	= "b_fname"
	f(26) 	= "b_lname"
	f(27) 	= "b_addr1"
	f(28) 	= "b_addr2"
	f(29) 	= "b_city"
	f(30) 	= "b_state"
	f(31) 	= "b_zip"
	f(32) 	= "b_country"
	f(33) 	= "b_phone"
	f(34) 	= "b_email"
	f(35) 	= "b_company"
	f(36) 	= "b_fax"
	f(37) 	= "b_web"
	f(38) 	= "comments"
	f(39) 	= "ship_method"
	f(40) 	= "order_weight"
	f(41) 	= "s_company"
	f(42) 	= "s_fname"
	f(43) 	= "s_lname"
	f(44) 	= "s_addr1"
	f(45) 	= "s_addr2"
	f(46) 	= "s_city"
	f(47) 	= "s_state"
	f(48) 	= "s_zip"
	f(49) 	= "s_country"
	f(50) 	= "s_phone"
	f(51) 	= "s_fax"
	f(52) 	= "s_web"
	f(53) 	= "s_email"
	f(54) 	= "s_addr_type"
	f(55) 	= "cust_track"
	f(56) 	= "cust_ip"
	f(57) 	= "item1"
	f(58) 	= "item2"
	f(59) 	= "item3"
	f(60) 	= "item4"
	f(61) 	= "item5"
	f(62) 	= "item6"
	f(63) 	= "item7"
	f(64) 	= "item8"
	
	qFields = ""
	
	For i = 1 to 64
		qFields = qFields & f(i) & ", "
		Select Case f(i)
		Case "cc_num", "cc_number", "cc_month", "cc_year"
			qReq = qReq & "'', "
		Case Else
			qReq = qReq & "'" & Request(f(i)) & "', "
		End Select
	Next
	
	qFields = left(qFields, len(qFields)-2)
	qReq = left(qReq, len(qReq)-2)
	
%>