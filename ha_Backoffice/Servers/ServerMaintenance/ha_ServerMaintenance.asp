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
	ColorTab = 1
	PageHeading = "Server Maintenance Log"

    Dim rs: Set rs = Server.CreateObject("ADODB.Recordset")

    If Request("action") = "delete" Then
      strSQL = "UPDATE ha_Server_Maintenance_Log SET DeletedDate = GETDATE(), DeletedBy = '" & Session("ha_shortname") & "' WHERE ID = " & Request("ID")
      connCasper10.execute strSQL
    End If

	OrderBy = ""
	If Request("OrderBy") <> "" Then
		OrderBy = Request("OrderBy")
		If OrderBy = "Assigned To" Then
			OrderBy = "AssignedTo"
		End If
		
		If Request("OrderBy") <> "" Then
		OrderBy = Request("OrderBy")
		If OrderBy = "Server" Then
			OrderBy = "CreatedDate"
		End If
		End if
		
		If Request("ad") = "DESC" Then
			OrderBy = OrderBy & " Desc"
		End If
	End If
	
	'Set up filtering data
	Call FilterCategories("Server", "Type", "", "", "")
	Call CheckBoxReadProperties(1, "Server Name")
	Call CheckBoxRead(2, "ServerMaintenanceType")
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
	var delTable = 'ha_Server_Maintenance_Log';
	
		
	<%=FunctionsFilter%>
	
	function clickAddMaintenanceEntry()
		{
        var URL = 'ha_ServerMaintenance_edit.asp' +
			'?tMode=Add';
        openPopUp(URL);
		}
		
    function Submit() 
		{
        document.frm.submit();
		}
	  				
	function clickSearch()
		{
		Submit();
		}
		
		
  </script>
  
  <script type="text/javascript">
  
		
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
				<%=CheckBoxFormat(1, "Server")%>
				<%=CheckBoxFormat(2, "Type")%>
			  </table>
			</td>
  		    <td width="84%">
			  <div class="ListBody">
				<%=TopRightLinksExpanded("", "", "", "", "", "", "", "", "", "Add Maintenance Entry")%>
				
				<table width="100%" cellpadding="2" cellspacing="2">
		  
				  <tr class="ListHeadingBar">
					<td width="3%" class="head" align="center">
					  Row
					</td>
					<td width="5%" class="head" align="center">
					  <%=ColumnHeading("Date")%>
					</td>
					<td width="10%" class="head">
					  <%=ColumnHeading("Server")%>
					</td>
					<td width="50%" class="head">
					  <table width="100%" border="0" cellpadding="0" cellspacing="0">
						<tr>
						  <td>
							<%=ColumnHeading("Task")%>
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
					<td width="10%" class="head">
					  <%=ColumnHeading("Type")%>
					</td>
					<td width="10%" class="head">
					  <%=ColumnHeading("Engineer")%>
					</td>
					<td width="15%" class="head" align="center">
					  Actions
					</td>
				  </tr>
		  
		  
<%
		strSQL = "ha_Servers_Maintenance_list('" & CheckBoxes(1,0) & "', '" & CheckBoxes(2,0) & "', " & _
		  "'" & OrderBy & "', '" & Request("SearchText") & "')"
		  
		rs.CursorLocation = 3
        rs.Open strSQL, connCasper10, 3, 3, 4		  
		iRecordperpage = 25
		
		'------------------------------------Starting Paging from Here-----------------------------------
%>
		<!--#include virtual="ha_backoffice/includes/FunctionsPageStart.inc"-->

				  <tr id="row[<%=row%>]" style="background-color: #FBFBFB;">
			  
					  <td class="row" style="<%=detailStyle%>" align="center">
						<%=row%>
					  </td>
					  <td class="row" style="<%=detailStyle%>" align="center">
						<%=rs("CreatedDate") %>
					  </td>
					  <td class="row" style="<%=detailStyle%>">
						<%=rs("ServerName")%>
					  </td>
					  <td class="row" style="<%=detailStyle%>">
						<%
						
							Response.Write iif(len(rs("Task")) > 61, mid(rs("Task"),1,60) & "...",rs("Task"))
						%>
					  </td>
					  <td class="row" style="<%=detailStyle%>">
						<%=rs("Type")%>
					  </td>
					  <td class="row" style="<%=detailStyle%>">
						<%=rs("Engineer")%>
					  </td>
					  <td class="row" style="<%=detailStyle%> text-align: center; font-size: 8pt;">
						<a href="javascript:void(0);" onclick="window.open('ha_ServerMaintenance_Edit.asp?tMode=Edit&ID=<%=rs("ID")%>','_blank','width=10,height=10')">Edit</a>
						<a href="javascript:clickDelete(<%=rs("ID")%>, '<%=replace(rs("Task"),"'","")%>', <%=row%>);">Delete</a>
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
