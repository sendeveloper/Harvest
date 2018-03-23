<%
'----------------------------------------------------------------------------------------------------------------------------		
  'Remove this include when  DB is finally moved.
  'Was placed here to prevent the removal of the connection when people edit the main connection include file and see a connection open
  'to an old server (Barley1)
  Dim objcon: Set objcon=server.CreateObject("ADODB.Connection")
  objcon.Open "driver=SQL Server;server=66.119.50.228,7843;uid=davewj2o;pwd=get2it;database=ha_prod"
'----------------------------------------------------------------------------------------------------------------------------
%>