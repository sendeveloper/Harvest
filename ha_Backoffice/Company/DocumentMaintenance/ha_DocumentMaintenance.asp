<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

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
	PageHeading = "Document Maintenance"

    Dim rs: Set rs = Server.CreateObject("ADODB.Recordset")

	OrderBy = ""
	If Request("OrderBy") <> "" Then
		Select Case Request("OrderBy")
		Case "Document Name"
			OrderBy = "FileName"
		Case "Type"
			OrderBy = "DocType"
		Case "Last Upload"
			OrderBy = "CreatedDate"
		Case Else
			OrderBy = Request("OrderBy")
		End Select
		If Request("ad") = "DESC" Then
			OrderBy = OrderBy & " Desc"
		End If
	End If
	
	'Set up filtering data
	Call FilterCategories("Type", "Business", "Department", "", "")
	Call CheckBoxRead(1, "DocMaint - Type")
	Call CheckBoxRead(2, "DocMaint - Business")
	Call CheckBoxRead(3, "DocMaint - Department")


%>

  <title>Harvest American Backoffice - <%=PageHeading%></title>
  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />

  <link href="<%=strPathIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet" />
  <link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>menu/dropdowntabfiles/halfmoontabs.css" />

  <script type="text/javascript" src="<%=strPathIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>
  <script type="text/javascript" src="<%=strPathBOIncludes%>haFunctions.js"></script>
  
  <script type="text/javascript">
	var delTable = 'ha_Documents';
  	<%=FunctionsFilter%>
	var EmpID = <%=Session("ha_EmpID")%>;
	var valueType;
	var valueBusiness;
	var valueDepartment;

    function clickAddDocument()
        {
		clickUpload(0);
		}
		
    function clickDownload(ID)
        {
        var URL = 'ha_Document_download.asp' +
            '?ID=' + ID;
        openPopUp(URL);
        }

	function clickEdit(ID)
        {
        var URL = 'ha_DocumentMaintenance_edit.asp' +
            '?ID=' + ID;
        openPopUp(URL);
        }
	  
    function clickInfo(ID)
        {
        var URL = 'ha_Document_info.asp' +
            '?ID=' + ID;
        openPopUp(URL);
        }
		
	function clickPreview(ID)
        {
        var URL = 'ha_Document_preview.asp' +
            '?ID=' + ID;
        openPopUp(URL);
        }
				
    function clickUpload(ID)
        {
        var URL = 'ha_Document_upload.asp' +
            '?ID=' + ID;
        openPopUp(URL);
        }
  </script>
  
  <style type="text/css">
  </style>
  
</head>

<body class="gray_desktop">
  <table align="center" cellspacing="0" cellpadding="0" class="MainBodyFixed">	
    <tr valign="top">
      <td>
        <!--#include virtual="ha_Backoffice/includes/bodyparts/ha_PageHeading.inc"-->
        <table class="MainTable">
          <form action="" method="post" name="frm">		

		  <tr valign="top">
		    <td width="16%">
			  <table class="filter">
			    <%=FilterHeading%>
				<%=CheckBoxFormat(1, "Type")%>				  
				<%=CheckBoxFormat(2, "Business")%>				  
				<%=CheckBoxFormat(3, "Department")%>				  
			  </table>
			</td>
  		    <td width="84%">
			  <div class="ListBody">
				<%=TopRightLinksExpanded("", "", "", "", "", "", "", "", "", "Add Document")%>

				<table width="100%" cellpadding="2" cellspacing="2">
				
				  <tr class="ListHeadingBar">
					<td width="5%" class="head" align="center">
					  Row
					</td>
					<td width="40%" class="head">
					  <table width="100%" border="0" cellpadding="0" cellspacing="0">
						<tr>
						  <td>
							<%=ColumnHeading("Document Name")%>
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
					<td width="10%" class="head" align="center">
					  <%=ColumnHeading("Business")%>
					</td>
					<td width="10%" class="head" align="center">
					  <%=ColumnHeading("Department")%>
					</td>
					<td width="35%" class="head" align="center">
					  Actions
					</td>
				  </tr>
		  
<%
		strSQL = "ha_Document_list('" & CheckBoxes(1,0) & "', '" & CheckBoxes(2,0) & "', '" & CheckBoxes(3,0) & "', " & _
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
<%
			fName = rs("fName")
			If len(fName) > 40 Then
				fName = left(fName,40) & " . . ."
			End If
%>
					<a href="javascript:clickDownload(<%=rs("ID")%>);" title="Click to Download"><%=fName%></a>
				  </td>
				  <td class="row" style="<%=detailStyle%>" align="center">
					<%=rs("Business")%>
				  </td>
				  <td class="row" style="<%=detailStyle%>" align="center">
					<%=rs("Department")%>
				  </td>
				  <td class="row" style="<%=detailStyle%> text-align: center; font-size: 8pt;">
					<a href="javascript:clickDownload(<%=rs("ID")%>);">Download</a>
					<a href="javascript:clickEdit(<%=rs("ID")%>);">Edit</a>
					<a href="javascript:clickInfo(<%=rs("ID")%>);">Info</a>
<%
	'If rs("ContentType") = "image/jpeg" Then
		Response.Write "<a href='javascript:clickPreview(" & rs("ID") & ");'>Preview</a>"
	'End If
%>
					<a href="javascript:clickUpload(<%=rs("ID")%>);">Upload</a>
					<a href="javascript:clickDelete(<%=rs("ID")%>, '<%=rs("fName")%>', <%=row%>);">Delete</a>
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

