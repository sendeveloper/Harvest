<!DOCTYPE html>

  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/Config.asp"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsFilter.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsPagination.inc"-->
  
<%
    Dim connString, objConn, strSQL
	ColorTab = 1
	PageHeading = "Database Classes"
        
    Dim rs: Set rs = Server.CreateObject("ADODB.Recordset")	
  
	OrderBy = ""
	If Request("OrderBy") <> "" Then
		Select Case Request("OrderBy")
		Case "Table Name"
			OrderBy = "TableName"
		Case Else
			OrderBy = Request("OrderBy")
		End Select
		If Request("ad") = "DESC" Then
			OrderBy = OrderBy & " Desc"
		End If
	End If
	
	'Set up filtering data
	Call FilterCategories("DatabaseClasses", "", "", "", "")
	CALL CheckBoxReadTableSimple(1, "ha_Database_Classes", "ClassName")	
%>

<html>
<head>

  <title>Harvest American Backoffice - <%=PageHeading%></title>
  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />

  <link href="<%=strPathIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet" />
  <link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>menu/dropdowntabfiles/halfmoontabs.css" />

  <script type="text/javascript" src="<%=strPathIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>
  <script type="text/javascript" src="<%=strPathBOIncludes%>haFunctions.js"></script>

  <script type="text/javascript">
	<%=FunctionsFilter%>
	
    function clickAddTable()
        {
        var URL = 'ha_Table_edit.asp' +
            '?ID=0';
        openPopUp(URL);
        }		
		
	function clickEditTable(ID)
        {
        var URL = 'ha_Table_edit.asp' +
            '?ID=' + ID;
        openPopUp(URL);
        }

	function clickTableReport()
        {
        var URL = 'ha_Table_Report.asp'
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
				<%=CheckBoxFormat(1, "Database Classes")%>				  
			  </table>
			</td>
  		    <td width="84%">
			  <div class="ListBody">
				<%=TopRightLinksExpanded("", "", "", "", "", "", "", "", "Table Report", "Add Table")%>
			
				<table width="100%" cellpadding="2" cellspacing="2">
		
				  <tr class="ListHeadingBar">
				  <td width="5%" class="head" align="center">
					Row
				  </td>
				  <td width="40%" class="head">
					<table width="100%" border="0" cellpadding="0" cellspacing="0">
					  <tr>
						<td>
						  <%=ColumnHeading("Table Name")%>
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
				  <td width="15%" class="head" align="center">
					Object Type
				  </td>
				  <td width="25%" class="head" align="center">
					Update Method
				  </td>
				  <td width="15%" class="head" align="center">
					Actions
				  </td>
				</tr>

<%
        Dim nRecordsPerPage, icurrentPage, intPageCount, intRecordCount           

		strSQL = "ha_DatabaseClasses_list('" & CheckBoxes(1,0) & "', '" & OrderBy & "', '" & Request("SearchText") & "')"

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
					  <%=rs("TableName")%>
					</td>
								
					<td class="row" style="<%=detailStyle%>" align="center">
					  <%=rs("ObjectType")%>
					</td>
					
					<td class="row" style="<%=detailStyle%>">
					  <%=rs("UpdateMethod")%>
					</td>
					
					<td class="row" style="<%=detailStyle%> text-align: center">
					  <a href="javascript:clickEditTable(<%=rs("ID")%>)">Edit</a>
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

