<!DOCTYPE html>
  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/Config.asp"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsDelete.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsFilter.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsPagination.inc"-->

<html>
<head>

  <%
	Session("Redirect") = ""
	ColorTab = 4
	PageHeading = "Employees"

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
	Call FilterCategories("Status", "Type", "", "", "")
	Call CheckBoxReadTable(1, "ha_EmployeesFormatted", "EmpStatus")
	Call CheckBoxReadTable(2, "ha_EmployeesFormatted", "EmpType")


%>

  <title>Harvest American Backoffice - <%=PageHeading%></title>
  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />

  <link href="<%=strPathIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet" />
  <link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>menu/dropdowntabfiles/halfmoontabs.css" />

  <script type="text/javascript" src="<%=strPathIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>
  <script type="text/javascript" src="<%=strPathBOIncludes%>haFunctions.js"></script>
  
  <script type="text/javascript">
	var delTable = 'ha_Employees';
  	<%=FunctionsFilter%>
	var EmpID = <%=Session("ha_EmpID")%>;

    function clickAddEmployee()
        {
        var URL = 'ha_Employee_edit.asp' +
            '?ID=0';
        openPopUp(URL);
        }
		
    function clickActiveList()
        {
        var URL = 'ha_Employee_List_active.asp';
        openPopUp(URL);
        }
		
	function clickEdit(ID)
        {
        var URL = 'ha_Employee_edit.asp' +
            '?ID=' + ID;
        openPopUp(URL);
        }
	  					
	function clickJobDescription(ID)
        {
        var URL = 'ha_EmployeeJobDescription.asp' +
            '?ID=' + ID;
        openPopUp(URL);
        }
		
	function clickPermissionsEdit(ID)
        {
        var URL = 'ha_EmployeePermissions_edit.asp' +
            '?ID=' + ID;
        openPopUp(URL);
        }
		
	function clickStatusEdit(ID)
        {
        var URL = 'ha_EmployeeStatus_edit.asp' +
            '?ID=' + ID;
        openPopUp(URL);
        }
		
<!--Function added to go with new code for number-it addition-->
	function clickPopupImage(ID)
		{
		var URL = '/ha_BackOffice/Company/DocumentMaintenance/ha_Document_Show.asp' +
			'?PhotoID=' + ID;
			PopupCenter(URL,'','900','500');
		}
<!--End of function that was added-->
		
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
		  
	  <!--Code added to make a hyperlink to the document to make a number-it account-->
	  	<tr>
			<td align="right" style="font-weight: bold; font-size: 10pt;" colspan="4">
			  <a href="javascript:clickPopupImage(1194);">Add Number-it Account</a>&nbsp;&nbsp;
			  <a href="javascript:clickPopupImage(1422);">Add Zip2Tax Account</a>
			</td>
		</tr>
		
	  <!--End of number-it code added-->
	  
		  <tr valign="top">
		    <td width="16%">
			  <table class="filter">
			    <%=FilterHeading%>
				<%=CheckBoxFormat(1, "Status")%>				  
				<%=CheckBoxFormat(2, "Type")%>				  
			  </table>
			</td>
  		    <td width="84%">

			  <div class="ListBody">
				<%=TopRightLinksExpanded("", "", "", "", "", "", "", "", "Active List", "Add Employee")%>
				
				<table width="100%" cellpadding="2" cellspacing="2">
		  
				  <tr class="ListHeadingBar">
				  <td width="5%" class="head" align="center">
					Row
				  </td>
				  <td width="35%" class="head">
					<table width="100%" border="0" cellpadding="0" cellspacing="0">
					  <tr>
						<td>
						  <%=ColumnHeading("Name")%>
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
					<%=ColumnHeading("Status")%>
				  </td>
				  <td width="15%" class="head" align="center">
					<%=ColumnHeading("Type")%>
				  </td>
				  <td width="30%" class="head" align="center">
					Actions
				  </td>
				</tr>
		  
		  
<%
		strSQL = "ha_Employees_list('" & CheckBoxes(1,0) & "', '" & CheckBoxes(2,0) & "', " & _
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
					<%=rs("FullName")%>
				  </td>
				  <td class="row" style="<%=detailStyle%> text-align: center">
					<%=rs("EmpStatus")%>
				  </td>
				  <td class="row" style="<%=detailStyle%> text-align: center">
					<%=rs("EmpType")%>
				  </td>
				  <td class="row" style="<%=detailStyle%> text-align: center">
					<span style="font-size: 8pt;">
					  <a href="javascript:clickEdit(<%=rs("ID")%>);">Edit</a>
					  <a href="javascript:clickJobDescription(<%=rs("ID")%>);">Job Description</a>
					  <a href="javascript:clickPermissionsEdit(<%=rs("ID")%>);">Permissions</a>
					  <a href="javascript:clickStatusEdit(<%=rs("ID")%>);">Status</a>
					  <a href="javascript:clickDelete(<%=rs("ID")%>, '<%=rs("FullName")%>', <%=row%>);">Delete</a>
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

