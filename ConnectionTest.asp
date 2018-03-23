<% 
	Response.Write("Hello.")

	
	
    Dim objcon: Set objcon=server.CreateObject("ADODB.Connection")
    objcon.Open "driver=SQL Server;server=68.178.202.54;uid=davewj2o;pwd=get2it;database=ha_prod" ' barley1

	
	
    Dim connAdmin: Set connAdmin=server.CreateObject("ADODB.Connection")
    connAdmin.Open "driver=SQL Server; server=208.109.189.101; uid=davewj2o; pwd=get2it; database=z2t_WebBackoffice" ' barley2

	
	
    Dim connPublic: Set connPublic=server.CreateObject("ADODB.Connection")
    connPublic.Open "driver=SQL Server; server=208.109.189.101; uid=davewj2o; pwd=get2it; database=z2t_WebPublic" ' barley2

		
	
    Dim connCasper: Set connCasper = server.CreateObject("ADODB.Connection")
    connCasper.Open "driver=SQL Server;server=66.119.55.118,7043;uid=davewj2o;pwd=get2it;database=z2t_BackOffice" ' Casper10
	
%>
