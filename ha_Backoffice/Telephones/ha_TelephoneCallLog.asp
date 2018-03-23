<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/Config.asp"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsDelete.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsFilter.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsPagination.inc"-->

  <!--#include virtual="includes/phone-functions.asp"-->
  
<%
	Session("Redirect") = ""
	ColorTab = 2
	PageHeading = "Telephone Call Log"

	requestReferrer = RequestParse(Request("Referrer"), Request.Servervariables("HTTP_REFERER"))
	
    Dim rs: Set rs = Server.CreateObject("ADODB.Recordset")		
	
	'Set up filtering data
	Call FilterCategories("Line", "Direction", "", "", "")
	Call CheckBoxRead(1, "Phone - Line")
	Call CheckBoxRead(2, "Phone - Direction")
%>

<html>
<head>
  <title>Harvest American Backoffice - <%=PageHeading%></title>
  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />

  <script language="JavaScript" src="includes/checkDate.js" type="text/javascript"></script>
  <script language="JavaScript" src="datepick/ts_picker.js" type="text/javascript"></script>
  <script language="JavaScript" src="<%=strPathIncludes%>haFunctions.js" type="text/javascript"></script>

  <link href="<%=strPathBOIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet" />
  <link href="<%=strPathControlPanel%>includes/HarvestAmerican.css" type="text/css" rel="stylesheet" />
  <link rel="stylesheet" type="text/css" href="<%=strPathBOIncludes%>menu/dropdowntabfiles/halfmoontabs.css" />
  <script type="text/javascript" src="<%=strPathBOIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>
  
  <script type="text/javascript">
  
	var http = new XMLHttpRequest();
	<%=FunctionsFilter%>

  </script>
  
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
				<%=CheckBoxFormat(1, "Line")%>				  
				<%=CheckBoxFormat(2, "Direction")%>				  
			  </table>
			</td>
  		    <td width="84%">
			  <table width="100%" cellpadding="2" cellspacing="2">
				<%=TopRightLinks(8, "", "", "", "")%>
		
				<tr style="background-color: #308F5A;">
				  <td width="5%" class="head" align="center">
					Row
				  </td>
				  <td width="15%" class="head" align="center">
					<%=ColumnHeading("Call Date/Time")%>
				  </td>
				  <td width="15%" class="head" align="center">
					<%=ColumnHeading("Extension")%>
				  </td>
				  <td width="15%" class="head" align="center">
					<%=ColumnHeading("Line")%>
				  </td>
				  <td width="15%" class="head" align="center">
					<%=ColumnHeading("Direction")%>
				  </td>
				  <td width="15%" class="head" align="center">
					<%=ColumnHeading("Phone Number")%>
				  </td>
				  <td width="15%" class="head" align="center">
					<%=ColumnHeading("Ring")%>
				  </td>
				  <td width="15%" class="head" align="center">
					<%=ColumnHeading("Duration")%>
				  </td>
				</tr>

<%
        Dim nRecordsPerPage, icurrentPage, intPageCount, intRecordCount           

		strSQL = "ha_Phone_list('" & CheckBoxes(1,0) & "', '" & CheckBoxes(2,0) & "', '" & OrderBy & "', '" & Request("SearchText") & "')"
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
				
					<td class="row" style="<%=detailStyle%>" align="center">
					  <%=rs("CallDate")%>
					</td>
					
					<td class="row" style="<%=detailStyle%>" align="center">
					  <%=rs("Ext")%>
					</td>
					
					<td class="row" style="<%=detailStyle%>" align="center">
					  <%=rs("CO")%>
					</td>
										
					<td class="row" style="<%=detailStyle%>" align="center">
					  <%=rs("Direction")%>
					</td>
									  
<%
	PhoneNumber = lTrim(rTrim(rs("PhoneNumber")))
	If len(PhoneNumber) = 10 Then
		PhoneNumber = Left(PhoneNumber,3) & "-" & Mid(PhoneNumber,4,3) & "-" & Right(PhoneNumber,4)
	End If
	If len(PhoneNumber) = 7 Then
		PhoneNumber = Left(PhoneNumber,3) & "-" & Right(PhoneNumber,4)
	End If
%>

					<td class="row" style="<%=detailStyle%>" align="right">
					  <%=PhoneNumber%>&nbsp;&nbsp;
					</td>
					
					<td class="row" style="<%=detailStyle%>" align="center">
					  <%=rs("Ring")%>
					</td>
					
					<td class="row" style="<%=detailStyle%>" align="center">
					  <%=rs("Duration")%>
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
