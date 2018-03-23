<html>
<head>
    <title>Harvest American Backoffice - Requirements Post</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />

</head>
<body>

<%
    Dim rs, strSQL

  	Set connLocal = Server.CreateObject("ADODB.Connection")
	connLocal.Open "driver=SQL Server; server=127.0.0.1,7043; uid=davewj2o; pwd=get2it; database=ha_PublishedTables"

    Set rs = server.createObject("ADODB.Recordset")
	
	strSQL = "SELECT * FROM ha_Products WHERE ID = " & Request("ID")
	rs.Open strSQL, connLocal, 2, 3
	rs("EditedBy") = Session("ha_shortname")
	rs("EditedDate") = Now
	
	rs("BreakoutHead1Name")	= Request("BreakoutHead1Name")
	rs("BreakoutHead1Seq")	= Request("BreakoutHead1Sequence")
	rs("BreakoutHead2Name")	= Request("BreakoutHead2Name")
	rs("BreakoutHead2Seq")	= Request("BreakoutHead2Sequence")
			
	rs("GLAccount")			= Request("GLAccountHarvest")
	rs("GLSequence")		= Request("GLSequenceHarvest")
	rs("GLAccountZip2Tax")	= Request("GLAccountZip2Tax")
	rs("GLSequenceZip2Tax")	= Request("GLSequenceZip2Tax")
	
    rs.Update 
    rs.Close
    connLocal.close
%>

<script type="text/javascript">
    //window.opener.location.href = window.opener.location.href;
    //window.close()
</script>

</body>
</html>
