<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/Config.asp"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsDelete.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsFilter.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsPagination.inc"-->
  
<%
    Dim connString, recordsPerPage, currentPage, objConn, strSQL, numOfPages
	ColorTab = 5
	PageHeading = "Domains"
        
    Dim rs: Set rs = Server.CreateObject("ADODB.Recordset")	
  
	OrderBy = ""
	If Request("OrderBy") <> "" Then
		Select Case Request("OrderBy")
		Case "Name"
			OrderBy = "DomainName"
		Case "Server"
			OrderBy = "ServerName"
		Case "Registrar"
			OrderBy = "DomainRegister"
		Case Else
			OrderBy = Request("OrderBy")
		End Select
		If Request("ad") = "DESC" Then
			OrderBy = OrderBy & " Desc"
		End If
	End If
	
	'Set up filtering data
	Call FilterCategories("Class", "Server", "Registrar", "", "")
	Call CheckBoxReadTable(1, "ha_Domains", "Class")
	Call CheckBoxReadTable(2, "ha_Domains", "ServerName")
	Call CheckBoxReadTable(3, "ha_Domains", "DomainRegister")
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
	var delTable = 'ha_Domains';
	<%=FunctionsFilter%>
	
    function clickAddDomain()
        {
        var URL = 'ha_Domains_edit.asp' +
            '?ID=0';
        openPopUp(URL);
        }		
		
    function clickDomainRead(url)
        {
        var URL = 'ha_Domain_read.asp?d=' + url
        openPopUp(URL);
        }
		
	function clickEditDomain(ID)
        {
        var URL = 'ha_Domains_edit.asp' +
            '?ID=' + ID;
        openPopUp(URL);
        }

	function clickWhoIs(domain)
        {
        var URL = 'http://who.godaddy.com/whoischeck.aspx' +
            '?domain=' + domain;
		window.open(URL,'','toolbar=yes, scrollbars=yes, resizable=yes, top=100, left=100, width=1000, height=800');			
        }
		
  </script>
  
  <style type="text/css">
  </style>
  
</head>

<body class="gray_desktop">
  <table align="center" cellspacing="0" cellpadding="0" class="MainBody">
    <tr>
      <td>
        <!--#include virtual="ha_Backoffice/includes/bodyparts/ha_PageHeading.inc"-->
        
        <form action="" method="post" name="frm">

        <table style="width: 1150px; margin: 0 auto;" cellpadding="0" cellspacing="0">
		  <tr valign="top">
		    <td width="16%">
			  <table class="filter">
			    <%=FilterHeading%>
				<%=CheckBoxFormat(1, "Class")%>				  
				<%=CheckBoxFormat(2, "Server")%>				  
				<%=CheckBoxFormat(3, "Registrar")%>				  
			  </table>
			</td>
  		    <td width="84%">
			  <table width="100%" cellpadding="2" cellspacing="2">
				<%=TopRightLinks(7, "", "", "", "Add Domain")%>
		
				<tr style="background-color: #308F5A;">
				  <td width="5%" class="head" align="center">
					Row
				  </td>
				  <td width="13%" class="head" align="center">
					<%=ColumnHeading("Subdomain")%>
				  </td>
				  <td width="23%" class="head" align="center">
					<%=ColumnHeading("Name")%>
				  </td>
				  <td width="13%" class="head" align="center">
					<%=ColumnHeading("Class")%>
				  </td>
				  <td width="13%" class="head" align="center">
					<%=ColumnHeading("Server")%>
				  </td>
				  <td width="13%" class="head" align="center">
					<%=ColumnHeading("Registrar")%>
				  </td>
				  <td width="20%" class="head" align="center">
					Actions
				  </td>
				</tr>

<%
        Dim nRecordsPerPage, icurrentPage, intPageCount, intRecordCount           

		strSQL = "ha_Domains_list('" & CheckBoxes(1,0) & "', '" & CheckBoxes(2,0) & "', '" & CheckBoxes(3,0) & "', '" & OrderBy & "', '" & Request("SearchText") & "')"
		'response.write strSQL

		rs.CursorLocation = 3
		rs.Open strSQL, connCasper10, 3, 3, 4
		iRecordperpage = 20
		
		'------------------------------------Starting Paging from Here-----------------------------------
%>
		<!--#include virtual="ha_backoffice/includes/FunctionsPageStart.inc"--> 

				  <tr id="row[<%=row%>]" style="background-color: #FBFBFB;">
				  
					<td class="row" style="<%=detailStyle%>" align="center">
					  <%=(iCurrentPage * iRecordperpage) + recordCount%>
					</td>
					
					<td class="row" style="<%=detailStyle%> text-align: right">
					  <%SubDomain = iif(isNullSpace(rs("SubName")),"",Null2Space(rs("SubName")) & ".")%>
					  <%=SubDomain%>
					</td>
					
					<td class="row" style="<%=detailStyle%>">
					  <%=rs("DomainName")%>
					</td>
					
					<td class="row" style="<%=detailStyle%> text-align: center">
					  <%=rs("Class")%>
					</td>
								
					<td class="row" style="<%=detailStyle%> text-align: center">
					  <%=rs("ServerName")%>
					</td>

					<td class="row" style="<%=detailStyle%> text-align: center">
					  <%=rs("DomainRegister")%>
					</td>

					<td class="row" style="<%=detailStyle%> text-align: center">

	                  <%
					    SubDomain = IIf(isnull(rs("SubDomain")),"",rs("SubDomain"))
					    url = IIf(rs("SecureSite") = "Yes", "https://", "http://")
					    url = url & IIf(SubDomain = "", "www.", SubDomain & ".")
						url = url & rs("DomainName")
					  %>

					  <a href="javascript:clickEditDomain(<%=rs("ID")%>);">Edit</a>
					  <a href="javascript:clickDelete(<%=rs("ID")%>, '<%=rs("DomainName")%>', <%=row%>);">Delete</a>
					  <a href="javascript:clickDomainRead('<%=url%>');">Read</a>
					  <a href="javascript:clickWhoIs('<%=rs("DomainName")%>');">WhoIs</a>
					</td>
				  </tr>

		<!--#include virtual="ha_backoffice/includes/FunctionsPageEnd.inc"-->

			  </table>
						
		  </td>
		</tr>
	  
		<%Call DisplayPageNavagation%>
		</table>
		</form>
		
	  </td>
	</tr>
		
	<tr><td>&nbsp;</td></tr>
  </table>
  <!--#include virtual="ha_BackOffice/includes/BodyParts/ha_PageTail.inc"-->
</body>
</html>

