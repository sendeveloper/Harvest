<%
		SQL =  "ha_GetCounts_ContactConstant_ExportTypes"
	   Dim connBackOffice
	   set connBackOffice=server.CreateObject("ADODB.Connection")
	   connBackOffice.Open "driver=SQL Server; server=66.119.50.228,7843; uid=davewj2o; pwd=get2it; database=ha_CRM"  'Casper08
		   Dim rs
	   set rs=server.createObject("ADODB.Recordset")
   rs.open SQL, connBackOffice, 2, 3

	   If Not rs.eof Then
	    	response.Write(rs("ha_Software")&"~"&rs("ha_Paper")&"~"&rs("ha_Both")&"~"&rs("z2t_TableUploadCustomer")&"~"&rs("z2t_FreeTrialIncomplete")&"~"&rs("z2t_FreeTrialComplete")&"~"&rs("z2t_CustomerRenewals")&"~"&rs("z2t_AllCustomers"))
		end if 

%>