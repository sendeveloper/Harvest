  <!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->
  
<html>
<head>
	<title>Harvest American Backoffice - Products Post</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />

</head>
<body>
<%
    Dim rs, strSQL

  	Set connLocal = Server.CreateObject("ADODB.Connection")
	connLocal.Open "driver=SQL Server; server=127.0.0.1,7043; uid=davewj2o; pwd=get2it; database=ha_PublishedTables"

    Set rs = server.createObject("ADODB.Recordset")
		
    If Request("ID") = "0" Then
      strSQL = "SELECT * FROM ha_Products"
      rs.Open strSQL, connLocal, 2, 3
      rs.AddNew
	  rs("CreatedBy") = Session("ha_shortname")
      rs("CreatedDate") = Now
    Else
      strSQL = "SELECT * FROM ha_Products WHERE ID = " & Request("ID")
      rs.Open strSQL, connLocal, 2, 3
	  rs("EditedBy") = Session("ha_shortname")
      rs("EditedDate") = Now
    End If
	
		rs("ActiveProduct") 			= Request("ActiveProduct")
		rs("BusinessCategory") 			= Request("BusinessCategory")
		rs("ProductCategory") 			= Request("ProductCategory")
		rs("ItemID") 					= iif(Request("ItemID")="", null, Request("ItemID"))
		rs("Sequence") 					= iif(Request("Sequence")="", null, Request("Sequence"))
		rs("Description") 				= Trim(Request("Description"))
		rs("ShoppingCartID")			= Request("ShoppingCartID")
		rs("DescriptionShoppingCart") 	= Trim(Request("DescriptionSC1"))
		rs("DescriptionShoppingCart2") 	= Trim(Request("DescriptionSC2"))
		rs("Color") 					= Request("Color")
		rs("Retail") 					= iif(Request("Retail")="", 0, Request("Retail"))
		rs("ImageName") 				= Request("ImageName")
		rs("IsInventoried") 			= Request("Inventoried")
		rs("InventoryName") 			= Request("InventoryName")
		rs("SKU") 						= Request("SKU")

    rs.Update 
    rs.Close
    connLocal.close
%>

<script type="text/javascript">
    window.opener.location.href = window.opener.location.href;
    window.close()
</script>

</body>	
</html>
