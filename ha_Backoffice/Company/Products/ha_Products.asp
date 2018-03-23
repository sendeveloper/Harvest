<!DOCTYPE html>
<html>
<head>
  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/Config.asp"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsDelete.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsFilter.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsPagination.inc"-->

<%
	ColorTab = 5
	PageHeading = "Products"

	Dim rs: Set rs = Server.CreateObject("ADODB.Recordset")	
  
	OrderBy = ""
	If Request("OrderBy") <> "" Then
		Select Case Request("OrderBy")
		Case "Seq"
			OrderBy = "Sequence"
		Case Else
			OrderBy = Request("OrderBy")
		End Select
		If Request("ad") = "DESC" Then
			OrderBy = OrderBy & " Desc"
		End If
	End If
	
	'Set up filtering data
	Call FilterCategories("BusinessCategory", "ProductCategory", "ItemIDRanges" , "", "")
	Call CheckBoxRead(1, "Product - BusinessCategory")
	Call CheckBoxRead(2, "Product - ProductCategory")
	Call CheckBoxRead(3, "Product - Range")
  
	Set connLocal = Server.CreateObject("ADODB.Connection")
	connLocal.Open "driver=SQL Server; server=127.0.0.1,7043; uid=davewj2o; pwd=get2it; database=ha_PublishedTables"

  If Request("action") = "dupl" Then
    strSQL = "SELECT * FROM ha_Products WHERE ProductID = " & Request("ID")
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open strSQL, connLocal, 1, 3

    Set insRS = Server.CreateObject("ADODB.Recordset")
    strSQL = "SELECT * FROM ha_Products"
    insRS.Open strSQL, connLocal, 1, 3
    insRS.AddNew

    insRS("ActiveProduct") = rs("ActiveProduct")
    insRS("CategoryID") = rs("CategoryID")
    insRS("ItemID") = rs("ItemID")
    insrs("Sequence") = rs("Sequence")
    insrs("InventoryName") = rs("InventoryName")
    insrs("Type") = rs("Type")
    insrs("Description") = rs("Description")
    insrs("DescriptionShoppingCart") = rs("DescriptionShoppingCart")
    insrs("DescriptionShoppingCart2") = rs("DescriptionShoppingCart2")
    insrs("Color") = rs("Color")
    insrs("Retail") = rs("Retail")
    insrs("ImageName") = rs("ImageName")
    If rs("IsInventoried") = "True" Then insrs("IsInventoried") = "1" Else insrs("IsInventoried") = "0"
    If rs("RequiresShipping") = "True" Then insrs("RequiresShipping") = "1" Else insrs("RequiresShipping") = "0"
    If rs("RequiresRegistration_rt") = "True" Then insrs("RequiresRegistration_rt") = "1" Else insrs("RequiresRegistration_rt") = "0"
    If rs("RequiresRegistration_nm") = "True" Then insrs("RequiresRegistration_nm") = "1" Else rs("RequiresRegistration_nm") = "0"
    If rs("RequiresRegistration_nmpro") = "True" Then insrs("RequiresRegistration_nmpro") = "1" Else insrs("RequiresRegistration_nmpro") = "0"
    If rs("RequiresRegistration_lookup") = "True" Then insrs("RequiresRegistration_lookup") = "1" Else insrs("RequiresRegistration_lookup") = "0"
    If rs("RequiresEmailTables") = "True" Then insrs("RequiresEmailTables") = "1" Else insrs("RequiresEmailTables") = "0"
    If rs("RequiresRegistration_initial") = "True" Then insrs("RequiresRegistration_initial") = "1" Else insrs("RequiresRegistration_initial") = "0"
    If rs("RequiresRegistration_updates") = "True" Then insrs("RequiresRegistration_updates") = "1" Else rs("RequiresRegistration_updates") = "0"
    If rs("RequiresRegistration_link") = "True" Then insrs("RequiresRegistration_link") = "1" Else insrs("RequiresRegistration_link") = "0"
    If rs("RequiresProcedureSetup") = "True" Then insrs("RequiresProcedureSetup") = "1" Else insrs("RequiresProcedureSetup") = "0"
    insrs("BreakoutHead1Name") = rs("BreakoutHead1Name")
    insrs("BreakoutHead1Seq") = rs("BreakoutHead1Seq")
    insrs("BreakoutHead2Name") = rs("BreakoutHead2Name")
    insrs("BreakoutHead2Seq") = rs("BreakoutHead2Seq")
    insrs("GLAccount") = rs("GLAccount")
    insrs("GLSequence") = rs("GLSequence")
    insrs("GLAccountZip2Tax") = rs("GLAccountZip2Tax")
    insrs("GLSequenceZip2Tax") = rs("GLSequenceZip2Tax")
    insrs("SKU") = rs("SKU")
    insrs.Update
    insrs.Close
    'connLocal.Execute strSQL
  End If
  
%>
  
  <title>Harvest American Backoffice - Products</title>
  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
  
  <link href="<%=strPathIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet">
  <link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>menu/dropdowntabfiles/halfmoontabs.css" media="screen" />
  
  <script type="text/javascript" src="<%=strPathIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>
  <script type="text/javascript" src="<%=strPathBOIncludes%>haFunctions.js"></script>
  
  <script type="text/javascript">
	var delTable = 'ha_PublishedTables.dbo.ha_Products';
	<%=FunctionsFilter%>

    function clickAddProduct()
        {
        var URL = 'ha_Products_edit.asp'
        openPopUp(URL);
        }
  </script>
  
  <style type="text/css">
  </style>
  
</head>

<body class="gray_desktop">
  <table class="MainBodyFixed">	
    <tr valign="top">
      <td>
        <!--#include virtual="ha_Backoffice/includes/bodyparts/ha_PageHeading.inc"-->
        <table class="MainTable">
          <form action="" method="post" name="frm">		
        
		  <tr valign="top">
		    <td width="16%">
			  <table class="filter">
			    <%=FilterHeading%>
				<%=CheckBoxFormat(1, "Business Category")%>			
				<%=CheckBoxFormat(2, "Product Category")%>					
				<%=CheckBoxFormat(3, "ItemID Ranges")%>				  
			  </table>
			</td>
  		    <td width="84%">
			  <div class="ListBody">
				<%=TopRightLinksExpanded("", "", "", "", "", "", "", "", "", "Add Product")%>
				
				<table width="100%" cellpadding="2" cellspacing="2">
		  
				  <tr class="ListHeadingBar">
				  <td width="5%" class="head" align="center">
					Row
				  </td>
				  <td width="8%" class="head" align="center">
					<%=ColumnHeading("Seq")%>
				  </td>
				  <td width="8%" class="head" align="center">
					<%=ColumnHeading("ItemID")%>
				  </td>
				  <td width="39%" class="head">
					<table width="100%" border="0" cellpadding="0" cellspacing="0">
					  <tr>
						<td>
						  <%=ColumnHeading("Description")%>
						</td>
						<td align="right">
						  <input type="text" id="searchText" name="searchText" value="">
						</td>
						<td align="left">
						  <a href="javascript:document.frm.submit();" class="buttonSearch">Search</a>
						</td>
					  </tr>
					</table>
				  </td>
				  <td width="8%" class="head" align="center">
					<%=ColumnHeading("Retail")%>
				  </td>
				  <td width="32%" class="head" align="center">
					Actions
				  </td>
				</tr>
			
<%
        Dim strSQL, nRecordsPerPage, icurrentPage, intPageCount, intRecordCount           
		strSQL = "ha_Product_list('" & CheckBoxes(1,0) &  "', '" & CheckBoxes(2,0) &  "', '" & CheckBoxes(3,0) & "', '" & OrderBy & "', '" & Request("SearchText") & "')"
		
		rs.CursorLocation = 3
		rs.Open strSQL, connLocal, 3, 3, 4
		iRecordperpage = 20
		
		'------------------------------------Starting Paging from Here-----------------------------------
%>
		<!--#include virtual="ha_backoffice/includes/FunctionsPageStart.inc"-->

				  <tr id="row[<%=row%>]" style="background-color: #FBFBFB;">
				  
					<td class="row" style="<%=detailStyle%>" align="center">
					  <%=(iCurrentPage * iRecordperpage) + recordCount%>
					</td>
					
					<td class="row" style="<%=detailStyle%>" align="center">
					  <%=rs("Sequence")%>
					</td>

					<td class="row" style="<%=detailStyle%>" align="center">
					  <%=rs("ItemID")%>
					</td>
					
					<td class="row" style="<%=detailStyle%>">
					  <%=rs("Description")%>
					</td>
								
					<td class="row" style="<%=detailStyle%>" align="right">
					  <%
					  Retail = iif(isNumeric(rs("Retail")),rs("Retail"),0)
					  strRetail = FormatNumber(Retail,2)
					  %>
					  <%=strRetail%>
					</td>
					
					<td class="row" style="<%=detailStyle%> text-align: center">
					  <span style="font-size: 8pt;">
						<a href="javascript:void(0)" onclick="openPopUp('ha_Products_edit.asp?id=<% =rs.Fields("id") %>')">Edit</a>
						<a href="javascript:void(0)" onclick="openPopUp('ha_Products_Requirements_edit.asp?id=<% =rs.Fields("id") %>')">Requirements</a>
						<a href="javascript:void(0)" onclick="openPopUp('ha_Products_Bookkeeping_edit.asp?id=<% =rs.Fields("id") %>')">Bookkeeping</a>
						<a href="javascript:void(0)" onclick="if(confirm('You are about to duplicate this record. Are you sure?')){window.location = 'ha_Products.asp?action=dupl&id=<% =rs.Fields("productid") %>';}">Duplicate</a>
					    <a href="javascript:clickDelete(<%=rs("ID")%>, '<%=replace(Null2Space(rs("Description")),"'", "\'")%>', <%=row%>);">Delete</a>
					  </span>
					</td>
				  </tr>

				<!--#include virtual="ha_backoffice/includes/FunctionsPageEnd.inc"-->
				</table>
			  </div>
			</td>
		  </tr>
		
		  <%Call DisplayPageNavagation%>
          </form>	
        </table>

	  </td>
	</tr>				  
  </table>
  <!--#include virtual="ha_BackOffice/includes/BodyParts/ha_PageTail.inc"-->
</body>
</html>
