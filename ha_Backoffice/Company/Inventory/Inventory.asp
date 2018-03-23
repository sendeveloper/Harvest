<!DOCTYPE html>
<!--#include virtual="includes/connection.asp"-->
<!--#include virtual="ha_BackOffice/includes/Config.asp"-->
<!--#include virtual="ha_backoffice/includes/FunctionsDelete.inc"-->
<!--#include virtual="ha_backoffice/includes/FunctionsFilter.inc"-->
<!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->
<!--#include virtual="ha_backoffice/includes/FunctionsPagination.inc"-->
<html>
    <head>
<%
	Session("Redirect") = ""
	ColorTab = 5
	PageHeading = "Inventory"

    Dim rs: Set rs = Server.CreateObject("ADODB.Recordset")

	OrderBy = ""
	If Request("OrderBy") = "" Then
		OrderBy = "Name"
	Else
		OrderBy = Request("OrderBy")
		If Request("ad") = "DESC" Then
			OrderBy = OrderBy & " Desc"
		End If
	End If
	'Set up filtering data
	'Call FilterCategories("Description", "AssignedTo", "", "", "")
	'Call CheckBoxReadTable(1, "ha_Inventory", "Description")
	'Call CheckBoxReadTable(1, "ha_Inventory", "AssignedTo")
%>

        <title>Harvest American Backoffice - <%=PageHeading%></title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    

	<link href="<%=strPathIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet" />
	<link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>menu/dropdowntabfiles/halfmoontabs.css" />


	<script type="text/javascript" src="<%=strPathIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>
    <script type="text/javascript" src="<%=strPathBOIncludes%>haFunctions.js"></script>

   <script type="text/javascript">
	var delTable = 'ha_Inventory';
  	<%=FunctionsFilter%>

    function clickAddToInventory()
        {
        var URL = 'BlankEdit.asp' +
            '?ID=0';
        openPopUp(URL);
        }
		
	function clickEdit(ID)
        {
        var URL = 'BlankEdit.asp' +
            '?ID=' + ID;
        openPopUp(URL);
        }
		
  </script>
  <style type="text/css">
  </style>
    </head>

<body class="gray_desktop">
  <table class="MainBodyFixed" cellpadding="10">
    <tr valign="top">
      <td>
        <!--#include virtual="ha_Backoffice/includes/bodyparts/ha_PageHeading.inc"-->
        <table class="MainTable">
          <form action="" method="post" name="frm">		
		  <tr valign="top">
		  <td width="16%">
			  <table class="filter">
			    <%=FilterHeading%>
				<%=CheckBoxFormat(1, "Description")%>				  			  
			  </table>
			</td>
			<td width="84%">
			  <div class="ListBody">
			<%=TopRightLinksExpanded("", "", "", "", "", "", "", "", "", "Add To Inventory")%>
			<table width="100%" cellpadding="2" cellspacing="2">
		  
				  <tr class="ListHeadingBar">
				  <td width="5%" class="head" align="center">
					Row
				  </td>
						<td width="15%" class="head" align="center">
						  <%=ColumnHeading("eType")%>
						</td>
						<td width="35%" class="head" align="center">
						  <%=ColumnHeading("Description")%>
						</td>
						<td width="15%" class="head" align="center">
						  <%=ColumnHeading("Condition")%>
						</td>
						<td width="15%" class="head" align="center">
						  <%=ColumnHeading("AssignedTo")%>
						</td>
						<td>
						  <%=ColumnHeading("Quantity")%>
						</td>
				  <td width="30%" class="head" align="center">
					Actions
				  </td>
				  <td width="60%" class="head">
					<table width="100%" border="0" cellpadding="0" cellspacing="0">
					  <tr>
						<td align="right">
						  <input type="text" id="searchText" name="searchText" value="">
						</td>
						<td align="left">
						  <a href="javascript:document.frm.submit();" class="buttonSearch">Search</a>
						</td>
					  </tr>
					</table>
				</td>
				</tr>
<%
		strSQL = "ha_Inventory_View"

		rs.CursorLocation = 3
		rs.Open strSQL, connCasper10, 3, 3, 4
		iRecordperpage = 25
		
		'------------------------------------Starting Paging from Here-----------------------------------
%>
<!--'" & CheckBoxes(1,0) & "', '" & CheckBoxes(2,0) & "', " & _
		  "'" & OrderBy & "', '" & Request("SearchText") & "'-->
		<!--#include virtual="ha_backoffice/includes/FunctionsPageStart.inc"--> 

				  <tr id="row[<%=row%>]" style="background-color: #FBFBFB;">
					<td class="row" style="<%=detailStyle%>" align="center">
					  <%=(iCurrentPage * iRecordperpage) + recordCount%>
					</td>
					<td class="row" style="<%=detailStyle%> text-align: center">
					  <%=rs("Type")%>
					</td>
					<td class="row" style="<%=detailStyle%> text-align: center">
					  <%=rs("Description")%>
					</td>
					<td class="row" style="<%=detailStyle%> text-align: center">
					  <%=rs("Condition")%>
					</td>
					<td class="row" style="<%=detailStyle%> text-align: center">
					  <%=rs("AssignedTo")%>
					</td>
					<td class="row" style="<%=detailStyle%>">
					  <%=rs("Quantity")%>
					</td>
					<td class="row" style="<%=detailStyle%> text-align: center">
					  <a href="javascript:clickEdit(<%=rs("ID")%>);">Edit</a>
					   <a href="javascript:clickDelete(<%=rs("ID")%>, '<%=rs("Type")%>', <%=row%>);">Delete</a>
				  </td>
				  </tr>
				<!--#include virtual="ha_backoffice/includes/FunctionsPageEnd.inc"-->
			  </table>
			  </div>
  <br><br><br><br>			  
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