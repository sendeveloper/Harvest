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
	ColorTab = 5
	PageHeading = "To Do List"

    Dim rs: Set rs = Server.CreateObject("ADODB.Recordset")

	OrderBy = ""
	If Request("OrderBy") <> "" Then
		OrderBy = Request("OrderBy")
		If OrderBy = "Assigned To" Then
			OrderBy = "AssignedTo"
		End If
		If Request("ad") = "DESC" Then
			OrderBy = OrderBy & " Desc"
		End If
	End If
	
	'Set up filtering data
	Call FilterCategories("Priority", "Assignee", "Category", "Status", "")
	Call CheckBoxRead(1, "ToDo - Priority")
	Call CheckBoxRead(2, "Assignee")
	Call CheckBoxRead(3, "ToDo - Category")
	Call CheckBoxRead(4, "ToDo - Status")

%>

  <title>Harvest American Backoffice - <%=PageHeading%></title>
  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />

  <link href="<%=strPathIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet" />
  <link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>menu/dropdowntabfiles/halfmoontabs.css" />

  <script type="text/javascript" src="<%=strPathIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>
  <script type="text/javascript" src="<%=strPathBOIncludes%>haFunctions.js"></script>
  
  <script type="text/javascript">
  
	var http = new XMLHttpRequest();
	var EmpID = <%=Session("ha_EmpID")%>;
	var delTable = 'ha_ToDo';
	<%=FunctionsFilter%>
	
	function clickAddToDoItem()
		{
        var URL = 'ha_ToDo_edit.asp' +
			'?tMode=Add';
        openPopUp(URL);
		}
		
    function Submit() 
		{
        document.frm.submit();
		}
	  		
	function clickFiles(ID)
        {
        var URL = 'ha_ToDo_Files.asp' +
            '?ID=' + ID;
        openPopUp(URL);
        }
		
	function clickCurrentProjects()
		{
        var URL = 'ha_ToDo_CurrentProjects.asp' +
			'?ID=' + <%=Session("ha_EmployeeID")%>;
        openPopUp(URL);
		}	
		
	function clickMeetingAgenda()
		{
		var URL = '/ha_BackOffice/Company/DocumentMaintenance/ha_Document_Show.asp' +
			'?PhotoID=1453';
		PopupCenter(URL,'','900','500');
        //openPopUp(URL);
		}
		
	function clickRules()
		{
        var URL = 'ha_ToDo_Rules.asp'
        openPopUp(URL);
		}
		
	function clickSearch()
		{
		Submit();
		}
		
	function clickStatistics()
		{
        var URL = 'ha_ToDo_Statistics.asp'
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
				<%=CheckBoxFormat(1, "Priority")%>				  
				<%=CheckBoxFormat(2, "Assignee")%>
				<%=CheckBoxFormat(3, "Category")%>
				<%=CheckBoxFormat(4, "Status")%>
			  </table>
			</td>
  		    <td width="84%">
			  <div class="ListBody">
				<%=TopRightLinksExpanded("", "", "", "", "", "Rules", "Meeting Agenda", "Current Projects", "Statistics", "Add To Do Item")%>

				<table width="100%" cellpadding="2" cellspacing="2">
		  
				  <tr class="ListHeadingBar">
					<td width="3%" class="head" align="center">
					  Row
					</td>
					<td width="5%" class="head" align="center">
					  <%=ColumnHeading("ID")%>
					</td>
					<td width="7%" class="head">
					  <%=ColumnHeading("Priority")%>
					</td>
					<td width="40%" class="head">
					  <table width="100%" border="0" cellpadding="0" cellspacing="0">
						<tr>
						  <td>
							<%=ColumnHeading("Title")%>
						  </td>
						  <td align="right">
							<input type="text" id="searchText" name="searchText" value="">
						  </td>
						  <td align="left">
							<a href="javascript:clickSearch();" class="buttonSearch">Search</a>
						  </td>
						</tr>
					  </table>
					</td>
					<td width="12%" class="head">
					  <%=ColumnHeading("Assigned To")%>
					</td>
					<td width="8%" class="head">
					  Category
					</td>
					<td width="10%" class="head">
					  <%=ColumnHeading("Status")%>
					</td>
					<td width="15%" class="head" align="center">
					  Actions
					</td>
				  </tr>
		  
		  
<%
		strSQL = "ha_ToDo_list('" & CheckBoxes(1,0) & "', '" & CheckBoxes(2,0) & "', '" & CheckBoxes(3,0) & "', '" & CheckBoxes(4,0) & "', " & _
		  "'" & OrderBy & "', '" & Request("SearchText") & "')"
		'response.write strSQL & "<br>"
		rs.CursorLocation = 3
        rs.Open strSQL, connCasper10, 3, 3, 4		  
		iRecordperpage = 30

		'------------------------------------Starting Paging from Here-----------------------------------
%>
		<!--#include virtual="ha_backoffice/includes/FunctionsPageStart.inc"-->

				  <tr id="row[<%=row%>]" style="background-color: #FBFBFB;">
			  
					  <td class="row" style="<%=detailStyle%>" align="center">
						<%=row%>
					  </td>
					  <td class="row" style="<%=detailStyle%>" align="center">
						<%=rs("ID") %>
					  </td>
					  <td class="row" style="<%=detailStyle%>" align="center">
						<%=rs("Priority") %>
					  </td>
					  <td class="row" style="<%=detailStyle%>">
						<%=rs("Title")%>
					  </td>
					  <td class="row" style="<%=detailStyle%>">
						<%=rs("AssignedTo")%>
					  </td>
					  <td class="row" style="<%=detailStyle%>">
						<%=rs("Category")%>
					  </td>
					  <td class="row" style="<%=detailStyle%>">
						<%=rs("Status")%>
					  </td>
					  <td class="row" style="<%=detailStyle%> text-align: center; font-size: 8pt;">
						<a href="javascript:void(0);" onclick="window.open('ha_ToDo_edit.asp?tMode=Edit&ID=<%=rs("ID")%>','_blank','width=10,height=10')">Edit</a>
						<a href="javascript:clickFiles(<%=rs("ID")%>);">Files<span style="color: #C0C0C0; text-decoration: none;">(<%=rs("SupportingFilesCount")%>)</span></a>
						<a href="javascript:clickDelete(<%=rs("ID")%>, '<%=replace(rs("Title"),"'","")%>', <%=row%>);">Delete</a>
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
