  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsPagination.inc"-->
  
<%
    Response.buffer=true
    Response.clear
	
	If Request("a") = "Attach" Then
		strSQL = "ha_Document_attach('ToDo', " & Request("ID") & ", " & Request("fid") & ", '" & Session("ha_shortname") & "')"
		connCasper10.Execute strSQL
	End If
	
	If Request("a") = "Detach" Then
		strSQL = "ha_Document_detach('ToDo', " & Request("ID") & ", " & Request("fid") & ", '" & Session("ha_shortname") & "')"
		connCasper10.Execute strSQL
	End If

	
    Title = "To Do - Supporting Files"

	Dim rs: Set rs = Server.CreateObject("ADODB.Recordset")
			
	strSQL = "SELECT * FROM ha_ToDo WHERE ID = " & Request("ID")
	rs.Open strSQL, connCasper10
		
	If Not rs.eof Then
		eDescription 	= rs("Description")
	End If
		
	rs.Close	
			
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

<html>
<head>
    <title><%=Title%></title>

    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <meta http-equiv="Cache-Control" content="no-cache">

    <link href="/ha_Backoffice/includes/ha_BackOfficePopup.css" type="text/css" rel="stylesheet">
    <script language="javascript" type="text/javascript" src="/includes/haFunctions.js"></script>

	<script type="text/javascript">
	    var httpFiles = new XMLHttpRequest();

		function clickClose()
			{
			window.opener.location.href = window.opener.location.href;
			window.close();
			}	
			
		function clickDownload(ID)
			{
			var URL = '/ha_BackOffice/Company/DocumentMaintenance/ha_Document_download.asp' +
				'?ID=' + ID;
			openPopUp(URL);
			}	

		function clickAttach()
			{
			document.getElementById('divEdit').style.visibility = "visible";
			}	
			
		function clickAttachClose()
			{
			document.getElementById('divEdit').style.visibility = "hidden";
			}	
			
		function clickAttachSelect(fid)
			{
			document.getElementById('divEdit').style.visibility = "hidden";
			var URL = 'ha_ToDo_files.asp?a=Attach&fid=' + fid + '&ID=<%=Request("ID")%>';
			window.location.href = URL;
			}	
			
		function clickDetach(fid)
			{
			var URL = 'ha_ToDo_files.asp?a=Detach&fid=' + fid + '&ID=<%=Request("ID")%>';
			window.location.href = URL;
			}	
			
		function clickSearch()
			{
			var f = document.getElementById('searchText').value;
			if (f > '')
				{
				getFiles(f);
				}
			}	
			
		function getFiles(f) 
			{		
			var url = '/ha_BackOffice/includes/ajax/ha_Documents_search.asp' +
				'?f=' + escape(f) +
				'&Now=' + escape(Date());
			
			httpFiles.open('GET', url, true);
			httpFiles.onreadystatechange = getFilesResponse;
			httpFiles.send();
			}

		function getFilesResponse() 
			{
			if (httpFiles.readyState == 4) 
				{
				if (httpFiles.status == 200) 
					{
					document.getElementById('divFiles').innerHTML = httpFiles.responseText;
					}
				}
			}
			
		function whichKey() 
			{
			if (event.keyCode==13)
				{
				if (document.getElementById('divEdit').style.visibility == "visible")
					{
					var f = document.getElementById('searchText').value;
					if (f > '')
						{
						clickSearch();
						}
					else
						{
						clickAttachClose();
						}
					}
				else
					{
					clickClose();
					}
				}
			}

	</script>

	<style type="text/css">
		td
			{
			font-size: 10pt;
			}

		td.label
			{
			font-weight: bold;
			text-align: right;
			}
			
		td.head
			{
			font-weight: bold;
			color: #C0C0C0;
			border-bottom: 1px solid #C0C0C0;
			text-align: center;
			}
			
		div.divEdit
			{
			position:absolute; 
			border: 2px solid black;
	        top: 100px; 
        	left: 200px; 
	        width: 450px; 
        	height: 450px;
	        background-color: #FFFFCC;
        	visibility: hidden;
			}
			
		a.fileRow			
			{
			text-decoration: none;
			color: #222222;
			}
			
		a.fileRow:hover
			{
			color: red;
			}
			
		td.fileRow
			{
			padding: 0;
			}
			
	</style
	
</head>


<body onLoad="SetScreen(900,700);" onkeydown="javascript:whichKey();">

  <span class="popupHeading"><%=Title%></span>

<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td>
	  
<%
		recordCount = 0
		LineCount = 0

		strSQL = "SELECT * FROM ha_ToDoFiles WHERE ToDoID = " & Request("id")
		rs.CursorLocation = 3
        rs.Open strSQL, connCasper10, 1, 3

		'------------------------------------Starting Paging from Here-----------------------------------
		iRecordperpage = 10
		rs.PageSize = iRecordperpage
		nRecordsPerPage = rs.PageSize
		rs.CacheSize = rs.PageSize
		intPageCount = rs.PageCount 
		intRecordCount = rs.RecordCount

		icurrentPage = TRIM(REQUEST("page"))
		if icurrentPage = "" then 
		  icurrentPage = 0
		else
		  icurrentPage = CInt(request("page"))
		end if

		If CInt(icurrentPage) > CInt(intPageCount) Then icurrentPage = intPageCount
		If CInt(icurrentPage) <= 0 Then intPage = 1
%>

	  <table style="width: 90%; margin: 0 auto;" cellpadding="2" cellspacing="2">
		<tr><td>&nbsp;</td></tr>
		
		<tr>
		  <td style="color: #C0C0C0; text-align: center;" colspan="5">
		    Here you can attach files from Document Maintenance to this To Do project.
		  </td>
		</tr>
		
		<tr><td>&nbsp;</td></tr>
		
		<tr valign="top">
		  <td style="text-align: center;" colspan="5">
			<span style="font-weight: bold;">To Do Project Name:</span>
		    <%=eDescription%>
		  </td>
		</tr>
		
		<tr><td>&nbsp;</td></tr>
		
		<tr>
		  <td align="right" colspan="5">
   		    <a href="javascript:clickAttach();">Attach New File</a>
		  </td>
		</tr>
		
		<tr><td>&nbsp;</td></tr>
		
		<tr>
		  <td width="5%" class="head">Row</td>
		  <td width="45%" class="head">Name</td>
		  <td width="15%" class="head">File Size</td>
		  <td width="15%" class="head">File Type</td>
		  <td width="20%" class="head">Actions</td>
		</tr>
		

<%		
        If intRecordCount > 0 Then
            rs.AbsolutePage = icurrentPage + 1
            intStart = rs.AbsolutePosition 
            If CInt(icurrentPage) = CInt(intPageCount)-1 Then
                intFinish = intRecordCount
            Else
                intFinish = intStart + (rs.PageSize - 1)
            End if


		  pageCount = 1
		  Do While pageCount <= iRecordperpage And Not rs.EOF
			recordCount = recordCount + 1
		  
		    rowStyle = "background-color: #FBFBFB;"
				
				If recordCount mod 3 = 0 Then
					detailStyle = "border-bottom: 3px solid #63C28D;"
				Else
					detailStyle = "border: 0;"
				End If
%>


		<tr>
		  <td align="center">
			<%=(iCurrentPage * iRecordperpage) + recordCount%>
		  </td>
		  <td>
		    <%=rs("fName")%>
		  </td>
		  <td align="right">
		    <%=FormatNumber(rs("fSize"),0)%>
		  </td>
		  <td align="center">
		    <%=rs("fType")%>
		  </td>
		  <td align="center">
   		    <a href="javascript:clickDetach(<%=rs("ImageID")%>);">Detach</a>
   		    <a href="javascript:clickDownload(<%=rs("ImageID")%>);">Download</a>
		  </td>
		</tr>
		
<%
                pageCount = pageCount + 1
                rs.MoveNext
              Loop
			
		End If
		rs.Close

		If recordCount > 0 Then
%>

		<tr>
		  <td style="border-top: 1px solid #C0C0C0;" colspan="6">
		    &nbsp;
		  </td>
		</tr>

<%
		End If
%>

			
		<tr>
		  <td colspan="6">
			<%
				'Vars = "&OrderBy=" & Request("OrderBy") & "&ad=" & Request("ad")
				Response.write GetHitCountAndPageLinks(Request.ServerVariables("URL"), intRecordCount, nRecordsPerPage, intPageCount, icurrentPage, "")
			%>
		  </td>
		</tr>
		
		<tr><td>&nbsp;</td></tr>
		
      </table>
		
    </td>
  </tr>
</table>

    <span class="popupButtons">
      <a href="javascript:clickClose();" class="bo_Button80">Close</a>
    </span>

<!-- In-page Popup Form -->
<div id="divEdit" name="divEdit" class="divEdit">
  <form action="javascript:whichKey();">
	<table width="100%" cellspacing="5" cellpadding="5">
	  <tr>
	    <td width="100%">
		  <span style="margin: auto; 
			font-weight: bold; 
			font-size: 16pt; 
			display: block;
			text-align: center;
			border-bottom: 1px solid black;">
			File Selection
		  </span>
		</td>
	  </tr>
	  
	  <tr>
		<td>
		  <input type="text" id="searchText" name="searchText" size="35" value="">
		  <a href="javascript:clickSearch();" class="bo_Button80">Search</a>
		</td>
	  <tr>
	  
	  <tr>
	    <td width="100%">
		  <div id="divFiles" name="divFiles"
			style="border: 1px solid black; 
		    height: 300px; 
		    background-color: white;
			overflow-y: scroll;">
		  
		  </div>
		</td>
	  </tr>
	    
	</table>
	
    <span class="popupButtons">
      <a href="javascript:clickAttachClose();" class="bo_Button80">Close</a>
    </span>
  </form>
 </div

</body>
</html>
