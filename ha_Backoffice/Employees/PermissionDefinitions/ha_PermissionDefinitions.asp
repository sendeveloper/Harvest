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
	Session("Redirect") = ""
	ColorTab = 4
	PageHeading = "Permission Definitions"

    Dim rs: Set rs = Server.CreateObject("ADODB.Recordset")

	OrderBy = ""
	If Request("OrderBy") <> "" Then
		OrderBy = Request("OrderBy")
		If Request("ad") = "DESC" Then
			OrderBy = OrderBy & " Desc"
		End If
	End If
	
	'Set up filtering data
	Call FilterCategories("Program", "Category", "", "", "")
	Call CheckBoxReadTable(1, "ha_PermissionDefinitions", "Program")
	Call CheckBoxReadTable(2, "ha_PermissionDefinitions", "Category")


%>

  <title>Harvest American Backoffice - <%=PageHeading%></title>
  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />

  <link href="<%=strPathIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet" />
  <link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>menu/dropdowntabfiles/halfmoontabs.css" />

  <script type="text/javascript" src="<%=strPathIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>
  <script type="text/javascript" src="<%=strPathBOIncludes%>haFunctions.js"></script>
  
  <script type="text/javascript">
	var delTable = 'ha_PermissionDefinitions';
  	<%=FunctionsFilter%>
	var EmpID = <%=Session("ha_EmpID")%>;
	
    function clickAddPermissionDefinition()
        {
		alert('Section not written yet');
        }
		
	function clickEdit(ID)
        {
        var URL = 'ha_DocumentMaintenance_edit.asp' +
            '?ID=' + ID;
        openPopUp(URL);
        }
	  
    function clickInfo(ID)
        {
        var URL = 'ha_PermissionDefinitions_info.asp' +
            '?ID=' + ID;
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
				<%=CheckBoxFormat(1, "Program")%>				  
				<%=CheckBoxFormat(2, "Category")%>				  
			  </table>
			</td>
  		    <td width="84%">
			  <div class="ListBody">
				<%=TopRightLinksExpanded("", "", "", "", "", "", "", "", "", "Add Permission Definition")%>
				
				<table width="100%" cellpadding="2" cellspacing="2">
		  
				  <tr class="ListHeadingBar">
					<td width="5%" class="head" align="center">
					  Row
					</td>
					<td width="10%" class="head">
					  <%=ColumnHeading("Program")%>
					</td>
					<td width="10%" class="head">
					  <%=ColumnHeading("Category")%>
					</td>
					<td width="10%" class="head">
					  <%=ColumnHeading("Page")%>
					</td>
					<td width="37%" class="head">
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
					<td width="18%" class="head" align="center">
					  Actions
					</td>
				  </tr>
		  
<%
		strSQL = "ha_PermissionDefinitions_list('" & CheckBoxes(1,0) & "', " & _
		  "'" & OrderBy & "', '" & Request("SearchText") & "')"
		  
		rs.CursorLocation = 3
        rs.Open strSQL, connCasper10, 3, 3, 4		  
		iRecordperpage = 25

		'------------------------------------Starting Paging from Here-----------------------------------
%>
		<!--#include virtual="ha_backoffice/includes/FunctionsPageStart.inc"-->

				  <tr id="row[<%=row%>]" style="background-color: #FBFBFB;">
				  
				  <td class="row" style="<%=detailStyle%>" align="center">
					<%=(iCurrentPage * iRecordperpage) + recordCount%>
				  </td>
				  <td class="row" style="<%=detailStyle%>">
					<%=rs("Program")%>
				  </td>
				  <td class="row" style="<%=detailStyle%>">
					<%=rs("Category")%>
				  </td>
				  <td class="row" style="<%=detailStyle%>">
					<%=rs("Page")%>
				  </td>
				  <td class="row" style="<%=detailStyle%>">
					<%=rs("Description")%>
				  </td>
				  <td class="row" style="<%=detailStyle%> text-align: center">
					<span style="font-size: 8pt;">
					  <a href="javascript:clickEdit(<%=rs("ID")%>);">Edit</a>
					  <a href="javascript:clickInfo(<%=rs("ID")%>);">Info</a>
					  <a href="javascript:clickDelete(<%=rs("ID")%>, '<%=rs("Description")%>', <%=row%>);">Delete</a>
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
	
    <tr><td>&nbsp;</td></tr>
  </table>
  <!--#include virtual="ha_BackOffice/includes/BodyParts/ha_PageTail.inc"-->
</body>
</html>

