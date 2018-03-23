<!DOCTYPE html>
  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/Config.asp"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsFilter.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsPagination.inc"-->
  
<%
	Session("Redirect") = ""
	ColorTab = 1
	PageHeading = "Table Distribution"

	OrderBy = RequestParse(Request("o"),"TableName")
		If Request("ad") = "DESC" Then
			OrderBy = OrderBy & " Desc"
		End If
	
    Dim rs: Set rs = Server.CreateObject("ADODB.Recordset")

	'OrderBy = ""
	'If Request("OrderBy") = "" Then
	'	OrderBy = "TableName"
	'Else
	'	OrderBy = Request("OrderBy")
	'	If Request("ad") = "DESC" Then
	'		OrderBy = OrderBy & " Desc"
	'	End If
	'End If
	
	'Set up filtering data
	Call FilterCategories("Category", "", "", "", "")
	Call CheckBoxReadTable(1, "ha_Table_Distribution_info", "Category")
%>

<html>
<head>
    <title>Harvest American Backoffice - <%=PageHeading%></title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
  
    <link href="<%=strPathIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet">
    <link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>menu/dropdowntabfiles/halfmoontabs.css">

    <script type="text/javascript" src="<%=strPathIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>
    <script type="text/javascript" src="<%=strPathBOIncludes%>haFunctions.js" language="javascript"></script>

    <style type="text/css">
		th
			{
			text-align:			left;
			border-bottom:		1px solid black;
			}

		th.Heading1
			{
			font-size:			14px;
			text-align:			center;
			border-bottom:		1px solid black;
			}

    </style>
	
	<script language="javascript" type = "text/javascript">
		<%=FunctionsFilter%>

        function clickPopup(n)
            {
            var URL = n + '.asp'
            openPopUp(URL);
            }
			
        function clickStatus(n)
            {
            var URL = '/ha_backoffice/Servers/TableDistribution/' + n
            openPopUp(URL);
            }
			
    </script>

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
			  </table>
			</td>
  		    <td width="84%">

			  <div class="ListBody">
				<%=TopRightLinksExpanded("", "", "", "", "", "", "", "", "", "")%>
				
				<table width="100%" cellpadding="2" cellspacing="2">
		  
				  <tr class="ListHeadingBar">
				  <td width="5%" class="head" align="center">
					Row
				  </td>
				  <td width="35%" class="head">
					<table width="100%" border="0" cellpadding="0" cellspacing="0">
					  <tr>
						<td>
						  <%=ColumnHeading("Table Name")%>
						</td>
						<td align="right">
						  <input type="text" id="searchText" name="searchText" value="" size="12">
						</td>
						<td align="left">
						  <a href="javascript:document.frm.submit();" class="buttonSearch">Search</a>
						</td>
					  </tr>
					</table>
				  </td>
				  <td width="15%" class="head" align="center">
					<%=ColumnHeading("Category")%>
				  </td>
				  <td width="15%" class="head" align="center">
					Frequency
				  </td>
				  <td width="15%" class="head" align="center">
					Last Run
				  </td>
				  <td width="15%" class="head" align="center">
					Actions
				  </td>
				</tr>
		  
		  
<%
		strSQL = "ha_Table_Distribution_list('" & CheckBoxes(1,0) & "', " & _
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
					<%=rs("TableName")%>
				  </td>
				  <td class="row" style="<%=detailStyle%> text-align: center">
					<%=rs("Category")%>
				  </td>
				  <td class="row" style="<%=detailStyle%> text-align: center">
					<%=rs("Frequency")%>
				  </td>
				  <td class="row" style="<%=detailStyle%> text-align: center">
					<%=rs("LastRunDuration")%>
				  </td>
				  <td class="row" style="<%=detailStyle%> text-align: center">
					<span style="font-size: 8pt;">
					
<%
	If rs("ActionDev") > "" Then
		ToolTip = ""
		If rs("ToolTipDev") > "" Then
			ToolTip = "<span class='ToolTipBox' style='text-align: left;'>" & rs("ToolTipDev") & "</span>"
		End If
%>
		<span class = 'ToolTip'><%=ToolTip%><a href="javascript:clickStatus('<%=rs("ActionDev")%>');">Dev</a></span>
<%
	End If
%>

<%
	If rs("ActionPromote") > "" Then
		ToolTip = ""
		If rs("ToolTipPromote") > "" Then
			ToolTip = "<span class='ToolTipBox' style='text-align: left;'>" & rs("ToolTipPromote") & "</span>"
		End If
%>
		<span class = 'ToolTip'><%=ToolTip%><a href="javascript:clickStatus('<%=rs("ActionPromote")%>');">Promote</a></span>
<%
	End If
%>

<%
	If rs("ActionProd") > "" Then
		ToolTip = ""
		If rs("ToolTipProd") > "" Then
			ToolTip = "<span class='ToolTipBox' style='text-align: left;'>" & rs("ToolTipProd") & "</span>"
		End If
%>
		<span class = 'ToolTip'><%=ToolTip%><a href="javascript:clickStatus('<%=rs("ActionProd")%>');">Prod</a></span>
<%
	End If
%>
					
<%
	If rs("ActionArchive") > "" Then
		ToolTip = ""
		If rs("ToolTipArchive") > "" Then
			ToolTip = "<span class='ToolTipBox' style='text-align: left;'>" & rs("ToolTipArchive") & "</span>"
		End If
%>
		<span class = 'ToolTip'><%=ToolTip%><a href="javascript:clickStatus('<%=rs("ActionArchive")%>');">Archive</a></span>
<%
	End If
%>
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


