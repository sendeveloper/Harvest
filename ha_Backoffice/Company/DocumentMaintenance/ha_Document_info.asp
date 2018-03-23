  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsPagination.inc"-->
  
<%
    Response.buffer=true
    Response.clear

    Title = "Document Information"

	Dim rs: Set rs = Server.CreateObject("ADODB.Recordset")
			
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
		function clickDownload(ID)
			{
			var URL = 'ha_Document_download.asp' +
				'?ID=' + ID;
			openPopUp(URL);
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
			
	</style
	
</head>


<body onLoad="SetScreen(900,700);">

  <span class="popupHeading"><%=Title%></span>

<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td>
	  
<%
		recordCount = 0
		LineCount = 0

		strSQL = "ha_Document_info(" & Request("id") & ")"
		rs.CursorLocation = 3
        rs.Open strSQL, connCasper10, 3, 3, 4		  

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
		
		<tr valign="top">
		  <td class="label" colspan="3">
			Document Name:
		  </td>
		  <td colspan="3">
		    <%=rs("fName")%>
		  </td>
		</tr>
		
		<tr valign="top">
		  <td class="label" colspan="3">
			Image ID:
		  </td>
		  <td colspan="3">
		    <%=rs("ID")%>
		  </td>
		</tr>
		
		<tr valign="top">
		  <td class="label" colspan="3">
			Type:
		  </td>
		  <td colspan="3">
		    <%=rs("DocType")%>
		  </td>
		</tr>

		<tr valign="top">
		  <td class="label" colspan="3">
			Business:
		  </td>
		  <td colspan="3">
		    <%=rs("Business")%>
		  </td>
		</tr>

		<tr valign="top">
		  <td class="label" colspan="3">
			Department:
		  </td>
		  <td colspan="3">
		    <%=rs("Department")%>
		  </td>
		</tr>
		
		<tr valign="top">
		  <td class="label" colspan="3">
			Originally Created By:
		  </td>
		  <td colspan="3">
		    <%=rs("OriginalCreatedBy")%>
		  </td>
		</tr>

		<tr valign="top">
		  <td class="label" colspan="3">
			Created Date:
		  </td>
		  <td colspan="3">
		    <%=rs("OriginalCreatedDate")%>
		  </td>
		</tr>

		<tr><td>&nbsp;</td></tr>
		
		<tr>
		  <td width="5%" class="head">Row</td>
		  <td width="15%" class="head">File Size</td>
		  <td width="20%" class="head">File Type</td>
		  <td width="15%" class="head">Uploaded By</td>
		  <td width="25%" class="head">Uploaded Date</td>
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
		  <td align="right">
		    <%=FormatNumber(rs("fSize"),0)%>
		  </td>
		  <td align="center">
		    <%=rs("fType")%>
		  </td>
		  <td align="center">
		    <%=rs("CreatedBy")%>
		  </td>
		  <td align="center">
		    <%=rs("CreatedDate")%>
		  </td>
		  <td align="center">
   		    <a href="javascript:clickDownload(<%=rs("ID")%>);">Download</a>
		  </td>
		</tr>
		
<%
                pageCount = pageCount + 1
                rs.MoveNext
              Loop
			
			End If
			rs.Close
%>

		
		<tr>
		  <td style="border-top: 1px solid #C0C0C0;" colspan="6">
		    &nbsp;
		  </td>
		</tr>
			
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
      <a href="javascript:window.close();" class="bo_Button80">Close</a>
    </span>

</body>
</html>
