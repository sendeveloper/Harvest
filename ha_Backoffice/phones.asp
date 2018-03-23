<%
	Session("Redirect") = ""
	ColorTab = 5
	PageHeading = "Harvest American Phone Call Log"
%>

<!--#include virtual="includes/connection.asp"-->
<!--#include virtual="includes/phone-functions.asp"-->

<html>
<head>
    <title>Harvest American Phone Call Log</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

    <script src="includes/checkDate.js" type="text/javascript"></script>
    <script src="datepick/ts_picker.js" type="text/javascript"></script>
    <script src="<%=strPathIncludes%>haFunctions.js"></script>

    <link href="<%=strPathControlPanel%>includes/HarvestAmerican.css" type="text/css" rel="stylesheet">
    <link href="<%=strPathBOIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet">
    <link rel="stylesheet" type="text/css" href="<%=strPathBOIncludes%>menu/dropdowntabfiles/halfmoontabs.css">
    <script type="text/javascript" src="<%=strPathBOIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>
<%
    referrer = Request("referrer")
    If referrer = "" Then
      referrer = Request.Servervariables("HTTP_REFERER")
    End If


%>
</head>

<body>
  <table align="center" cellspacing="0" cellpadding="0" class="MainBody">
    <tr>
	  <td>

		<!--#include virtual="ha_Backoffice/includes/bodyparts/ha_PageHeading.inc"-->
		
		<table width="98%" align="center" border="0" cellspacing="0" cellpadding="0">

		  <tr> 
			<td width="80%" align="left"> 
			</td>
			<td width="20%" height="30" align="right">
			  <a href="phone-statistics.asp" class="button">Statistics</a>
			  <a href="<%=referrer%>" class="button">Back</a>
			</td>
		  </tr>

		  <tr>
			<td colspan="2">
		<%
				  recordCount = 0
				  LineCount = 0

				  set rs=server.createObject("ADODB.Recordset")
				  rs.open "ha_Phone_CallLog_view", objcon, 3, 3, 4

		'------------------------------------Starting Paging from Here-----------------------------------
			iRecordperpage = 25
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
		<%      If intRecordCount > 0 Then %>
			  <table width="100%" align="left" cellspacing="0" cellpadding="2" class="table";>
				<tr bgcolor="#FFFF66">
				  <th style="width: 1em; max-width: 1em; overflow: hide;">#</th>
				  <th style="width: 8em;">Call Date</th>
				  <th style="width: 10em;">Extension</th>
				  <th style="width: 4em">Line</th>
				  <th style="width: 1em; max-width: 1em;">In</th>
				  <th style="width: 5em;">Dial Number</th>
				  <th style="width: 4em;">Ring</th>
				  <th style="width: 8em;">Duration</th>
				  <!-- <th style="width: 11em">Cost</th> -->
				  <!-- <th style="width: 10em">Access Code</th> -->
				  <!-- <th style="width: 3em">CD</th> -->
				  
				  <!-- <th style="width: 50%"></th> -->
				</tr>
		<%
					rs.AbsolutePage = icurrentPage + 1
					intStart = rs.AbsolutePosition 
					If CInt(icurrentPage) = CInt(intPageCount)-1 Then
						intFinish = intRecordCount
					Else
						intFinish = intStart + (rs.PageSize - 1)
					End if

					n = 1
					  Do While n <= iRecordperpage
						'on error resume next
							 recordCount = recordCount + 1
						  LineCount = LineCount + 1
						  If LineCount > 2 then
							LineCount = 0
							strColor = "#FFFFCC"
						  Else
							strColor = "#FFFFFF"
						  End If

							If rs.eof Then 
		%>
				<tr><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td></tr>
		<%                      
							Else
		%>

				<tr bgcolor=<%=strColor%>>
				  <td style="border-color: #EEEEEE;">
					<%=irecordperpage * icurrentpage + recordCount%>
				  </td>
				  <td style="border-color: #EEEEEE;">
					<%=rs("CallDate")%>
				  </td>
				  <td style="border-color: #EEEEEE;">
					<%=rs("Ext")%>
				  </td>
				  <td style="border-color: #EEEEEE;">
					<%=rs("CO")%>
				  </td>
				  <td style="border-color: #EEEEEE;">
					<% If rs("Incoming") = "True" Then %>
					  I
					<% End If %>
				  </td>
				  <td style="border-color: #EEEEEE;">
					<% If rs("Incoming") = "True" Then %>
					  <span style="font-weight: bold">
					<% End If %>
		<%          If Trim(rs("DialNumber")) = "" Then %>
					  &lt;Unknown&gt;
		<%         ElseIf IsNumeric(rs("DialNumber")) And len(trim(cstr(rs("DialNumber")))) = 10 Then
					 Response.Write("(" & mid(rs("DialNumber"), 1, 3) & ") " & mid(rs("DialNumber"), 4, 3) & "-" & mid(rs("DialNumber"), 7, 10))
				   Else %>
					 <%=rs("DialNumber")%>
		<%         End If %>
					<% If rs("Incoming") = "True" Then %>
					  </span>
					<% End If %>

				  </td>
				  <td style="border-color: #EEEEEE;">
					  <%=rs("Ring")%>
				  </td>
				  <td style="border-color: #EEEEEE;">
					<%=rs("Duration")%>
				  </td>
				  <!--
				  <td style="border-color: #EEEEEE;">
					<%=rs("Cost")%>
				  </td>
				  -->
				  <!--
				  <td style="border-color: #EEEEEE;">
					<%=rs("AccCode")%>
				  </td>
				  -->
				  <!--
				  <td style="border-color: #EEEEEE;">
					<%=rs("CD")%>
				  </td>
				  -->
				  <!-- 
				  <td style="border-color: #EEEEEE;">
				  </td> 
				  -->
				</tr>
		<%
							rs.MoveNext
							End If
							n = n + 1
						 Loop

		'        End If
				rs.close
				set rs = nothing
				objcon.close
		%>

				<tr bgcolor="#FFFF66">
				  <td>
					<b>Calls:</b>&nbsp;
					<%=iRecordPerPage * iCurrentPage + 1%>&nbsp;-&nbsp;<%=iRecordPerPage * iCurrentPage + recordCount%>&nbsp;/&nbsp;<%=intrecordCount%>
				  </td>
				  <td></td>
				  <td></td>
				  <td></td>
				  <td></td>
				  <td></td>
				  <td></td>
				  <td></td>
				</tr>
			  </table>
		<%
				Else
			'on error resume next
		%>
				  <td align="center" style="color: C0C0C0; font-weight: bold;">No records found matching that search criteria</td>
		<%      End If %>
		  <%=GetHitCountAndPageLinks(Request.ServerVariables("URL"), intRecordCount, nRecordsPerPage, intPageCount, icurrentPage, "&amp;referrer=" & referrer)%>
			</td>
			<td>

		<!-- end paging -->
			</td>
		  </tr>
		  
		</table>
	  </td>
	</tr>
    <tr>
      <td>
        <div style="margin-left: 5em; height: 4em"><!--#include virtual="ha_BackOffice/includes/BodyParts/copyright.asp"--></div>
      </td>
    </tr>
  </table>
</body>
</html>
