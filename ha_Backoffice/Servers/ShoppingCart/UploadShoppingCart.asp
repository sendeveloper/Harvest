<%
	Response.buffer=true
	Dim rs
	Dim SQL
	
    Set rs = server.createObject("ADODB.Recordset")
%>


<!--#include virtual="ha_backoffice/Servers/TableDistribution/includes/TableDistribution_TaskLog_write.inc"-->

<!--#include file="ShoppingCart_FieldList.inc"-->

<%
	'Response.write qFields & vbCrLf
	'Response.write qReq & vbCrLf
%>

<!--#include file="ShoppingCart_WriteToCasper08.inc"-->
<!--#include file="ShoppingCart_SendTextNotifications.inc"-->
<!--#include file="ShoppingCart_SendNotificationToGoogle.inc"-->
<!--#include file="ShoppingCart_SendNotificationToYotpo.inc"-->


<%
	'--Sample URL string for testing
	'/ha_BackOffice/Servers/ShoppingCart/UploadShoppingCart.asp
		'?inv_num=1234&date=03%2F01%2F2006&time=12%3A15%3A37&subtotal=100.00&ship_charge=5.00&sales_tax=7.00&cod_charge=0.00
		'&min_charge=0.00&vol_discount=0.00&pro_discount=0.00&total=112.00&pay_method=Credit%20Card&b_fname=John&b_lname=Doe
		'&b_addr1=123%20Test%20St.&b_addr2=&b_city=Testville&b_state=FL&b_zip=32837&b_country=United%20States&b_phone=321-555-1234
		'&b_email=test%40test123.com&b_company=&b_fax=&b_web=&comments=I%20love%20your%20products!	
%>

