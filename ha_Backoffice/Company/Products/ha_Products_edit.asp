<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsHelp.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->
  
<%
  	Set connLocal = Server.CreateObject("ADODB.Connection")
	connLocal.Open "driver=SQL Server; server=127.0.0.1,7043; uid=davewj2o; pwd=get2it; database=ha_PublishedTables"
	Dim rs: Set rs = Server.CreateObject("ADODB.Recordset")

	If IsNull(Request("ID")) or Request("ID")=0 Then
		ID = 0
		Title = "Product Add"
		
		eInventoried = 0
	Else
		ID = Request("ID")
		Title = "Product Edit"
		
		strSQL = "SELECT * FROM ha_Products WHERE ID = " & ID
        rs.open strSQL, connLocal, 0, 1

	    If Not rs.EOF Then
			eItemID				= rs("ItemID")
			eDescription 		= rs("Description")
			
			eActiveProduct 		= rs("ActiveProduct")
			eCategory 			= rs("CategoryID")
			eProductCategory 	= rs("ProductCategory")
			eItemID 			= rs("ItemID")
			eSequence 			= rs("Sequence")
			eBusinessCategory	= rs("BusinessCategory")
			eDescription 		= rs("Description")
			eShoppingCartID		= rs("ShoppingCartID")
			eDescriptionSC1 	= rs("DescriptionShoppingCart")
			eDescriptionSC2 	= rs("DescriptionShoppingCart2")
			eColor 				= rs("Color")
			eRetail 			= rs("Retail")
			eRetail				= iif(eRetail > 0, FormatNumber(eRetail,2), FormatNumber(0,2))
			eImageName 			= rs("ImageName")
			eInventoried		= rs("IsInventoried")
			eInventoryName		= rs("InventoryName")
			eSKU 				= rs("SKU")
			
			'eBC = rs("BusinessCategory")
			'eTY = rs("Type")
			
			eCreatedBy			= Null2Space(rs("CreatedBy"))
			eCreatedDate		= Null2Space(rs("CreatedDate"))
			eEditedBy			= Null2Space(rs("EditedBy"))
			eEditedDate			= Null2Space(rs("EditedDate"))
		End If
		
        rs.close
	End If
%>

<html>
<head>

  <title>Harvest American Backoffice - <%=Title%></title>
  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
  
  <link href="/ha_BackOffice/includes/ha_BackofficePopup.css" type="text/css" rel="stylesheet">
  <script type="text/javaScript" src="/ha_BackOffice/includes/haFunctions.js"></script> 
  
  <style type="text/css">    
    input[type="text"]
		{
		width: 12em;   
		}
				
	select
		{
		width: 12.4em;   
		}
	
	textarea
		{
		font-family:Verdana, Arial, Helvetica, sans-serif;
		font-size: 	10pt;
		width: 		400px;
		height:		3em;
		resize: 	none;
		}
  </style>
</head>

<body onload="SetScreen(850,850)">
  <form action="ha_Products_post.asp?ID=<%=ID%>" method="post" name="frm">

	<span class="popupHeading"><%=Title%></span>

    <table width="95%" border="0" cellspacing="2" cellpadding="2">

		<tr><td>&nbsp;</td></tr>
		
		<tr>
		  <td width="30%" align="right">
			Item ID:
		  </td>
		  <td width="70%" style="color: #1F3A4A; font-weight: bold;">
			<%=eItemID%>
		  </td>		 
		</tr>
		
		<tr>
		  <td align="right">
			Description:
		  </td>
		  <td style="color: #1F3A4A; font-weight: bold;">
			<%=eDescription%>
		  </td>		 
		</tr>

		<tr><td>&nbsp;</td></tr>
			
		<tr>
		  <td align="right">
			Active Product:
		  </td>
		  <td>
			<Select id="ActiveProduct" name="ActiveProduct">
				
<%
				strSQL = "ha_BackOffice.dbo.util_HTML_option_list('ha_PublishedTables.dbo.ha_types', 'class', 'TrueFalse', 'value', 'description', 'sequence', '', '" & eActiveProduct & "')"
			
				rs.Open strSQL, connLocal, 3, 3, 4
				If not rs.eof Then
					While not rs.eof
						Response.write rs("result")
						rs.MoveNext
					Wend
				End If
				rs.Close
%>
		
			</Select>
			<%=PopupHelp(2)%>
		  </td>		 
		</tr>
		
		<tr>
		  <td align="right">
			Item ID:
		  </td>
		  <td>
			<input type="text" name="ItemID" id="ItemID"
				size="10" value="<%=eItemID%>">
			<%=PopupHelp(18)%>
			<Span class="inputSuffix">(Always 3 digits)</span>
		  </td>		 
		</tr>

		<tr>
		  <td align="right">
			Sequence:
		  </td>
		  <td>
			<input type="text" name="Sequence" id="Sequence"
				size="10" value="<%=eSequence%>">
		  </td>		 
		</tr>
		
		<tr>
		  <td align="right">
			Business Category:
		  </td>
		  <td>
			<Select id="BusinessCategory" name="BusinessCategory">
				
<%
				strSQL = "ha_BackOffice.dbo.util_HTML_option_list('ha_PublishedTables.dbo.ha_types', 'class', 'Product - BusinessCategory', 'value', 'description', 'sequence', '', '" & eBusinessCategory & "')"
	
				rs.Open strSQL, connLocal, 3, 3, 4
				If not rs.eof Then
					While not rs.eof
						Response.write rs("result")
						rs.MoveNext
					Wend
				End If
				rs.Close
%>
		
			</Select>
			<%=PopupHelp(4)%>
		  </td>		 
		</tr>
		<tr>
		
		  <td align="right">
			Product Category:
		  </td>
		  <td>
			<Select id="ProductCategory" name="ProductCategory">
				
<%
				strSQL = "ha_BackOffice.dbo.util_HTML_option_list('ha_PublishedTables.dbo.ha_types', 'class', 'Product - ProductCategory', 'value', 'description', 'sequence', '', '" & eProductCategory & "')"
			
				rs.Open strSQL, connLocal, 3, 3, 4
				If not rs.eof Then
					While not rs.eof
						Response.write rs("result")
						rs.MoveNext
					Wend
				End If
				rs.Close
%>
		
			</Select>
			<%=PopupHelp(33)%>
		  </td>		 
		</tr>
		<tr>
		  <td align="right" vAlign="top">
			Description:
		  </td>
		  <td>
			<textarea id="Description" name="Description"><%=eDescription%></textarea>
		  </td>		 
		</tr>
		
		<tr>
		  <td align="right">
			Shopping Cart ID:
		  </td>
		  <td>
			<input type="text" name="ShoppingCartID" id="ShoppingCartID"
				size="10" value="<%=eShoppingCartID%>">
			<%=PopupHelp(9)%>
		  </td>		 
		</tr>
		
		<tr>
		  <td align="right" vAlign="top">
			Shopping Cart Description:
		  </td>
		  <td>
			<textarea id="DescriptionSC1" name="DescriptionSC1"><%=eDescriptionSC1%></textarea>
			<%=PopupHelp(5)%>
		  </td>		 
		</tr>
		
		<tr>
		  <td align="right" vAlign="top">
			Shopping Cart 2 Description:
		  </td>
		  <td>
			<textarea id="DescriptionSC2" name="DescriptionSC2"><%=eDescriptionSC2%></textarea>
		  </td>		 
		</tr>
		
		<tr>
		  <td align="right">
			Color/Details:
		  </td>
		  <td>
			<input type="text" name="Color" id="Color"
				size="10" value="<%=eColor%>">
			<%=PopupHelp(6)%>
		  </td>		 
		</tr>

		<tr>
		  <td align="right">
			Price:
		  </td>
		  <td>
			<input type="text" name="Retail" id="Retail"
				size="10" value="<%=eRetail%>">
		  <span class="inputSuffix">(Numbers and one period only, no commas or dollar signs)</span>
		  </td>		 
		</tr>

		<tr>
		  <td align="right">
			Image Name:
		  </td>
		  <td>
			<input type="text" name="ImageName" id="ImageName"
				size="10" value="<%=eImageName%>">
		  </td>		 
		</tr>
		
		<tr>
		  <td align="right">
			Is Inventoried:
		  </td>
		  <td>
			<Select id="Inventoried" name="Inventoried">
				
<%
				strSQL = "ha_BackOffice.dbo.util_HTML_option_list('ha_PublishedTables.dbo.ha_types', 'class', 'TrueFalse', 'value', 'description', 'sequence', '', '" & eInventoried & "')"
			
				rs.Open strSQL, connLocal, 3, 3, 4
				If not rs.eof Then
					While not rs.eof
						Response.write rs("result")
						rs.MoveNext
					Wend
				End If
				rs.Close
%>
		
			</Select>
		  </td>		 
		</tr>
		
		<tr>
		  <td align="right">
			Inventory Name:
		  </td>
		  <td>
			<input type="text" name="InventoryName" id="InventoryName"
				size="10" value="<%=eInventoryName%>">
		  </td>		 
		</tr>

		<tr>
		  <td align="right">
			SKU Number:
		  </td>
		  <td>
			<input type="text" name="SKU" id="SKU"
				size="10" value="<%=eSKU%>">
			<%=PopupHelp(7)%>
		  </td>		 
		</tr>

		<tr><td>&nbsp;</td></tr>
		
    </table>
		
		
    <span class="popupCredit"><%=CreditLine(eCreatedBy, eCreatedDate, eEditedBy, eEditedDate)%></span>	

    <span class="popupButtons">
      <a href="javascript:document.frm.submit();" class="bo_Button80">Submit</a>
      <a href="javascript:window.close();" class="bo_Button80">Close</a>
    </span>	
	
  </form>
</body>
</html>
