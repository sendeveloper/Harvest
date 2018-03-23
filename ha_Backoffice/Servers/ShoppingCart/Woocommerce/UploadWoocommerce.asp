<%
	Response.buffer=true
	Dim rs
	Dim SQL
	
    Set rs = server.createObject("ADODB.Recordset")
%>


<!--#include virtual="ha_backoffice/Servers/TableDistribution/includes/TableDistribution_TaskLog_write.inc"-->

<!--#include file="ShoppingCart_Woocommerce_WriteToCasper08.inc"-->


