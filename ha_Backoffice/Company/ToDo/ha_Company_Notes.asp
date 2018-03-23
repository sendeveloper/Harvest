<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/Config.asp"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsDelete.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsFilter.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsPagination.inc"-->
    <!--#Include virtual="ha_backoffice/includes/FunctionsNotes.inc"-->
  
<%
		
if Request.QueryString("a") = "Edit" Then
		strSQL = "ha_Note_write " & Request("noteid") & ",'" &  FixAps(REPLACE(Request("Note"),"'","&apos;")) & "','"  & Session("ha_shortname") &  "','" & Request("noteid") & "','" + Request("Category") + "'"
		'Response.Write  strSql
		'Response.end
		connLocal.Execute strSQL
		
		'response.redirect "ha_Company_Notes.asp?ID=" & Request("ID")
End If

    Dim connString, objConn, strSQL
	ColorTab = 5
	PageHeading = "Company Notes"
        
    Dim rs: Set rs = Server.CreateObject("ADODB.Recordset")	
  
	OrderBy = ""
	If Request("OrderBy") <> "" Then
		Select Case Request("OrderBy")
		Case "Company Name / Website"
			OrderBy = "CompanyName"
		Case Else
			OrderBy = Request("OrderBy")
		End Select
		If Request("ad") = "DESC" Then
			OrderBy = OrderBy & " Desc"
		End If
	End If
	
	'Set up filtering data
	Call FilterCategories("Category", "", "", "", "")
	Call CheckBoxRead(1, "ToDo - Category")
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
	var delTable = 'ha_Accomplishments';
	<%=FunctionsFilter%>
	var NoteID='';
    function clickAddANote()
        {
        var URL = 'ha_Note_edit.asp' +
            '?ID=0';
        openPopUp(URL);
        }		
		
function clickEditSave()
{
			var note = document.getElementById('Note').value;
			var URL = 'ha_Company_Notes.asp?noteid=' + NoteID + '&Note=' + note + '&a=Edit&Category='+Cat;
			//alert(URL);
			window.location.href = URL;
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
			
			
				<table width="100%" cellpadding="2" cellspacing="2">
		
				  <tr class="ListHeadingBar">
				  <td width="3%" class="head" align="center">
					Row
				  </td>
				  
				  <td width="7%" class="head" align="center">
					<%=ColumnHeading("Category")%>
				  </td>
				  <td width="60%" class="head">
					<table width="100%" border="0" cellpadding="0" cellspacing="0">
					  <tr>
						<td>
						  <%=ColumnHeading("Note")%>
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
				  <td  class="head"  width="20%">
						  <%=ColumnHeading("Date")%>
						</td>
				  <td width="10%" class="head" align="center">
					 Actions
				  </td>
				</tr>

<%
        Dim nRecordsPerPage, icurrentPage, intPageCount, intRecordCount, strDate,OrderBy2  
        strDate = instr(1,OrderBy,"Date")
        if strDate > 0 then
			OrderBy2 = OrderBy
			OrderBy = replace(OrderBy2,"Date","CreatedDate")
		end if
		strSQL = "ha_Notes_list(" & "'" & CheckBoxes(1,0) & "','','" & OrderBy & "','" & Request("SearchText") & "')"
		'response.write strSql
		'response.end
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
					
					<td class="row" style="<%=detailStyle%> text-align: center">
					  <%=rs("Category")%>
					</td>
					
					<td class="row" style="<%=detailStyle%>">
					  <% response.write iif(len(rs("Note")) > 61, mid(rs("Note"),1,60) & "...",rs("Note")) %>
					</td>
					<td class="row" style="<%=detailStyle%>">
					  <%=rs("CreatedDate")%>
					</td>
								
					<td class="row" style="<%=detailStyle%> text-align: center">
					<%if Session("ha_shortname") = rs("CreatedBy") then%>
					<%
					dim strNote
					strNote = Replace(rs("Note"),"&apos;","~")
					
					%>
					  <a href="javascript:clickNoteEdit2(<%=rs("ID")%>,'<%=Replace(Replace(rs("Note"),"&apos;","~"),"'","~")%>','<%=rs("Category")%>');">Edit</a>
					<%else%>
					<a href="javascript:clickViewNote(<%=rs("ID")%>,'<%=Replace(Replace(rs("Note"),"'","~"),"&apos;","~")%>');">View</a>
					<%End if%>
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

