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
	PageHeading = "Employee Calendar"

    Dim rs: Set rs = Server.CreateObject("ADODB.Recordset")
	
	OrderBy = ""
	If Request("OrderBy") <> "" Then
		OrderBy = Request("OrderBy")
		If Request("ad") = "DESC" Then
			OrderBy = OrderBy & " Desc"
		End If
	End If
	
	'Set up filtering data
	Call FilterCategories("Types", "Status", "", "", "")
	Call CheckBoxRead(1, "EmployeeCalendarTypes")
	Call CheckBoxRead(2, "EmployeeCalendarStatus")


%>

  <title>Harvest American Backoffice - <%=PageHeading%></title>
  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />

  <link href="<%=strPathIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet" />
  <link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>menu/dropdowntabfiles/halfmoontabs.css" />

  <script type="text/javascript" src="<%=strPathIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>
  <script type="text/javascript" src="<%=strPathBOIncludes%>haFunctions.js"></script>
  
  <script type="text/javascript">
	var delTable = 'ha_EmpCalendar';
  	<%=FunctionsFilter%>
	var EmpID = <%=Session("ha_EmpID")%>;

    function clickAddCalendarItem()
        {
        var URL = 'ha_EmployeeCalendar_edit.asp' +
            '?ID=0';
        openPopUp(URL);
        }
		
    function clickCalendar()
        {
        var URL = 'ha_Calendar.asp'
        openPopUp(URL);
        }
				
	function clickEdit(ID)
        {
        var URL = 'ha_EmployeeCalendar_edit.asp' +
            '?ID=' + ID;
        openPopUp(URL);
        }
		
	function clickTimeOffProcedure()
		{
		var URL = '/ha_BackOffice/Company/DocumentMaintenance/ha_Document_Show.asp' +
			'?PhotoID=1503';
			PopupCenter(URL,'','900','500');
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
				<%=CheckBoxFormat(1, "Types")%>				  
				<%=CheckBoxFormat(2, "Status")%>				  
			  </table>
			</td>
  		    <td width="84%">
			  <div class="ListBody">
				<%=TopRightLinksExpanded("", "", "", "", "", "", "", "Calendar", "Time Off Procedure", "Add Calendar Item")%>
				
				<table width="100%" cellpadding="2" cellspacing="2">
		  
				  <tr class="ListHeadingBar">
				  <td width="5%" class="head" align="center">
					Row
				  </td>
				  
				  <td width="12%" class="head" align="center">
					<%=ColumnHeading("Date")%>
				  </td>
				  
				  <td width="47%" class="head">
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
				  
				  <td width="12%" class="head" align="center">
					<%=ColumnHeading("Type")%>
				  </td>
				  
				  <td width="12%" class="head" align="center">
					<%=ColumnHeading("Status")%>
				  </td>
				  <td width="12%" class="head" align="center">
					Actions
				  </td>
				</tr>
		  
		  
<%
		strSQL = "ha_Employee_Calendar_list('" & CheckBoxes(1,0) & "', '" & CheckBoxes(2,0) & "', " & _
		  "'" & OrderBy & "', '" & Request("SearchText") & "')"

		rs.CursorLocation = 3
        rs.Open strSQL, connCasper10, 3, 3, 4		  
		iRecordperpage = 25

		'------------------------------------Starting Paging from Here-----------------------------------
%>
		<!--#include virtual="ha_backoffice/includes/FunctionsPageStart.inc"-->

<%
		Description	= rs("Description")
		If Len(Description) > 50 Then
			Description = Left(Description,50) & " . . ."
		End If
%>
		
				<tr id="row[<%=row%>]" style="background-color: #FBFBFB;">
				  <td class="row" style="<%=detailStyle%>" align="center">
					<%=row%>
				  </td>
				  <td class="row" style="<%=detailStyle%> text-align: center">
					<%=rs("StartDate")%>
				  </td>
				  <td class="row" style="<%=detailStyle%>">
					<%=Description%>
				  </td>
				  <td class="row" style="<%=detailStyle%> text-align: center">
					<%=rs("Type")%>
				  </td>
				  <td class="row" style="<%=detailStyle%> text-align: center">
					<%=rs("Status")%>
				  </td>
				  <td class="row" style="<%=detailStyle%> text-align: center">
					<span style="font-size: 8pt;">
					  <a href="javascript:clickEdit(<%=rs("ID")%>);">Edit</a>
					  <a href="javascript:clickDelete(<%=rs("ID")%>, '<%=replace(rs("Description"),"'","")%>', <%=row%>);">Delete</a>
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

