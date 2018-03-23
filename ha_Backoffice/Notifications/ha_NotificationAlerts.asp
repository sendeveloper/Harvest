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
	ColorTab = 9
	PageHeading = "Notification Alerts"

    Dim rs: Set rs = Server.CreateObject("ADODB.Recordset")

    If Request("action") = "delete" Then
      strSQL = "UPDATE ha_NotificationAlerts SET DeletedDate = GETDATE(), DeletedBy = '" & Session("ha_shortname") & "' WHERE ID = " & Request("ID")
      connCasper10.execute strSQL
    End If

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
	Call FilterCategories("Category", "Status", "", "", "")
	Call CheckBoxRead(1, "NotificationCategory")
	Call CheckBoxRead(2, "NotificationStatus")

%>

  <title>Harvest American Backoffice - <%=PageHeading%></title>
  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />

  <link href="<%=strPathIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet" />
  <link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>menu/dropdowntabfiles/halfmoontabs.css" />

  <script type="text/javascript" src="<%=strPathIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>
  <script type="text/javascript" src="<%=strPathBOIncludes%>haFunctions.js"></script>
  
  <script type="text/javascript">
  
	var EmpID = <%=Session("ha_EmpID")%>;
	var delTable = 'ha_NotificationAlerts';
	<%=FunctionsFilter%>
	
	function clickAddNotificationAlert()
		{
		alert('Here we need a typical Edit/Post pair of pop-ups.');
		/*
        var URL = 'ha_NotificationAlert_edit.asp' +
			'?rec=0';
        openPopUp(URL);
		*/
		}
		
	function clickEditNotificationAlert(rec)
		{
        var URL = 'ha_NotificationAlert_edit.asp' +
			'?rec=' + rec;
        openPopUp(URL);
		}
		
	function clickEditAlerts(rec)
		{
		alert('Here is where we would edit what to do in the case of a failed condition.');
		}
		
	function clickEditText(rec)
		{
		alert('Here is where we would edit the text that gets sent for both mobile or E-mail regarding this alert.');
		}
		
	function clickFlowchart()
		{
        var URL = 'ha_NotificationAlerts_FlowChart.asp'
        openPopUp(URL);
		}
		
	function clickSearch()
		{
		Submit();
		}
		
    function Submit() 
		{
        document.frm.submit();
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
				<%=CheckBoxFormat(1, "Category")%>
				<%=CheckBoxFormat(2, "Status")%>
			  </table>
			</td>
  		    <td width="84%">
			  <div class="ListBody">
				<%=TopRightLinksExpanded("", "", "", "", "", "", "", "", "Flowchart", "Add Notification Alert")%>
				
				<table width="100%" cellpadding="2" cellspacing="2">
		  
				  <tr class="ListHeadingBar">
					<td width="3%" class="head" align="center">
					  Row
					</td>
					<td width="5%" class="head" align="center">
					  <%=ColumnHeading("ID")%>
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
					<td width="8%" class="head">
					  <%=ColumnHeading("Category")%>
					</td>
					<td width="10%" class="head">
					  <%=ColumnHeading("Status")%>
					</td>
					<td width="15%" class="head" align="center">
					  Actions
					</td>
				  </tr>
		  
		  
<%
		strSQL = "ha_NotificationAlerts_list('" & CheckBoxes(1,0) & "', '" & CheckBoxes(2,0) & "', " & _
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
					<%=rs("ID") %>
				  </td>
				  <td class="row" style="<%=detailStyle%>">
					<%=rs("Title")%>
				  </td>
				  <td class="row" style="<%=detailStyle%>">
					<%=rs("Category")%>
				  </td>
				  <td class="row" style="<%=detailStyle%>">
					<%=rs("Status")%>
				  </td>
				  <td class="row" style="<%=detailStyle%> text-align: center; font-size: 8pt;">
					<a href="javascript:void(0);" onclick="window.open('ha_NotificationsAlerts_edit.asp?ID=<%=rs("ID")%>','_blank','width=10,height=10')">Alert</a>
					<a href="javascript:clickEditText(<%=rs("ID")%>);">Text</a>
					<a href="javascript:clickEditNotificationAlert(<%=rs("ID")%>);">Edit</a>
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
