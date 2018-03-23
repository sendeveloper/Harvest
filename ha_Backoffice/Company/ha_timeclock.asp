<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <!--#include virtual="includes/connection.asp"-->
    <!--#include virtual="ha_backoffice/includes/Config.asp"-->
    <!--#include file="sql.asp"-->
    <!--#include file="lib.asp"-->
    <%
    Dim rs, strsql, eName, eEmpID, eDOB, eHire, eTerm, eYearW, eVacAll, eVacUse, ePTOAll, ePTOUse, monthAnn, dayAnn, URLTC, qString

    URLTC = "/timeclock/"

    ' check for delete query
    If Not Request("delete") = "" Then
      connCasper10.Execute "DELETE FROM ha_EmpHours WHERE ID = " & Request("delete")
    End If
    'Response.write Request("id")
    If Not Request("id") = "" Then
      strsql = "SELECT * FROM ha_EmpAccounts WHERE ID ='" & Request("id") & "'"
      Set rs = Server.CreateObject("ADODB.Recordset")
      rs.Open strsql, connCasper10, 1, 3
      'response.write strsql
      eName = rs("FirstName")
      eEmpID = rs("ID")
      eDOB = rs("DOB")
      'eYearW = rs("YearsWCo")
      eHire = CDate(rs("HiredDate"))
      qString = "&id=" & Request("id")
      If Not rs("TerminationDate") = "" Then eTerm = rs("TerminationDate")
      rs.Close
      Set rs = Nothing
    Else
      strsql = "SELECT * FROM ha_EmpAccounts WHERE [UserName] ='" & Session("ha_Login") & "'"
      Set rs = Server.CreateObject("ADODB.Recordset")
      rs.Open strsql, connCasper10, 1, 3
      eName = rs("FirstName")
      eEmpID = rs("ID")
      eDOB = rs("DOB")
      'eYearW = rs("YearsWCo")
      eHire = CDate(rs("HiredDate"))
      If Not rs("TerminationDate") = "" Then eTerm = rs("TerminationDate")
      rs.Close
      qString = ""
      Set rs = Nothing
    End If

    monthAnn = Month(eHire)
    dayAnn = Day(eHire)

    
    Session("Redirect") = strPathIncludes & "login.asp"
	ColorTab = 5
	PageHeading = eName & "'s Time Clock"

    Function CheckNumber(num)
      If num MOD 3 = 0 Then
        Response.Write " style=""background-color: #B4D7BF"""
      End If
    End Function
    %>
    <title>Harvest American - <%=PageHeading%></title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <link href="http://code.jquery.com/ui/1.10.2/themes/south-street/jquery-ui.css" rel="stylesheet" />
    <link href="<%=strPathIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet" />
    <link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>menu/dropdowntabfiles/halfmoontabs.css" />
    <script type="text/javascript" src="<%=strPathIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>
    <script type="text/javascript">
        function popup(url) {
            var width = 600;
            var height = 425;
            var left = (screen.width - width) / 2;
            var top = (screen.height - height) / 2;
            var params = 'width=' + width + ', height=' + height;
            params += ', top=' + top + ', left=' + left;
            params += ', directories=no';
            params += ', location=no';
            params += ', menubar=no';
            params += ', resizable=no';
            params += ', scrollbars=yes';
            params += ', status=no';
            params += ', toolbar=no';
            newwin = window.open(url, 'windowname5', params);
            if (window.focus) { newwin.focus() }
            return false;
        }
    </script>
    <style type="text/css">
        tr.calrowhead
        {
            background-color: #008000;
        }
    </style>
</head>
<body class="gray_desktop">
  <table cellpadding="0" cellspacing="0" class="MainBody" style="margin: 1em auto 0; width: 1200px">
    <tr>
      <td>
        <!--#include virtual="ha_Backoffice/includes/bodyparts/ha_PageHeading.inc"-->
      </td>
    </tr>
    <tr>
      <td>
        <% 
          strsql = "ha_TimeClockStatistics '" & eEmpID & "'"
          'response.write strsql & chr(13)
          Set rs = Server.CreateObject("ADODB.Recordset")
          rs.Open strsql, connCasper10, 1, 3 %>
        <table style="width: 80%; margin: 0 auto" cellpadding="5" cellspacing="5">
          <tr class="calrowhead">
            <th style="color: white; width: 33%;">Hire Date</th>
            <th style="color: white; width: 33%;">Unpaid Hours</th>
            <th style="color: white; width: 33%;">Hours Worked YTD</th>
          </tr>
          <tr>
            <td style="text-align: center"><%=monthAnn & "/" & dayAnn%></td>
            <td style="text-align: center"><%If Not rs("UnpaidHours") = "" Then Response.write rs("UnpaidHours") Else Response.write "0" %></td>
            <td style="text-align: center"><%=rs("HoursTotal")%></td>
          </tr>
        </table>
        <table style="width: 80%; margin: 0 auto" cellpadding="5" cellspacing="5">
          <tr class="calrowhead">
            <th style="color: white; width: 25%;">Vacation Used</th>
            <th style="color: white; width: 25%;">Vacation Allowed</th>
            <th style="color: white; width: 25%;">Paid Time Off Used</th>
            <th style="color: white; width: 25%;">Paid Time Off Allowed</th>
          </tr>
          <tr>
            <td>
            <%
              Response.write rs("VacationUsed") 
            %>
            </td>
            <td>
            <%
              Response.write rs("VacationTotal")
            %>
            </td>
            <td>
            <%
              Response.write rs("PaidTimeOffUsed")
            %>
            </td>
            <td>
            <%
              Response.write rs("PaidTimeOffTotal")
            %>
            </td>
          </tr>
          <% rs.close: Set rs = Nothing %>
          <tr>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td style="text-align: center" colspan="2"><a href="javascript: void(0)" onclick="popup('<%=URLTC%>ha_timeclockEdit.asp?cMode=Add<%=iif(Request("id")= "", "", "&empid=" & eEmpID) %>')" class="button">Add Record</a></td>
          </tr>
        </table>
        <br /><br />
        <%
          Dim recordsPerPage, currentPage, numOfPages, intRecordCount, count: count = 1

          recordsPerPage = 20
          
          strsql = "SELECT * FROM ha_EmpHours WHERE [EmpID] = " & eEmpID & " ORDER BY [DateWorked] DESC"
          Set rs = Server.CreateObject("ADODB.Recordset")
          'response.Write strsql
          rs.Open strsql, connCasper10, 1, 3

          If Not rs.EOF Then
          rs.PageSize = recordsPerPage
	      numOfPages = rs.PageCount
          
          ' check for page
	      If Request.QueryString("Page") = "" Then
		    currentPage = 1
            rs.AbsolutePage = 1
	      Else
		    currentPage = CINT(Request.QueryString("Page")) ' convert Page request to an integer
            rs.AbsolutePage = CINT(Request.QueryString("Page")) ' convert Page request to an integer
	      End If

          intRecordCount = rs.RecordCount
          
        %>
        <table style="width: 80%; margin: 0 auto" cellpadding="5" cellspacing="5">
          <tr class="calrowhead">
            <th style="color: white; width: 100px">Date</th>
            <th style="color: white; width: 80px">Hours Worked</th>
            <th style="color: white; width: 100px">Type of Hours</th>
            <th style="color: white; width: 100px">Date Paid</th>
            <th style="color: white; width: 450px">Notes</th>
            <th style="color: white; width: 130px">Options</th>
          </tr>
        <%
          Do While Not count > 20 And Not rs.EOF
            Response.write "          <tr" & CheckNumber(count) & ">" & chr(13)
            Response.write "            <td>" & rs("DateWorked") & "</td>" & chr(13)
            Response.write "            <td style=""text-align: center"">" & rs("QuantityHours")& "</td>" & chr(13)
            Response.write "            <td>"
            Select Case rs("HoursType")
              Case 1
                response.write("Regular")
              Case 2
                response.write("Vacation")
              Case 3
                response.write("Personal")
              Case 4
                response.write("Holiday")
            End Select
            Response.write "            </td>" & chr(13)
            Response.write "            <td>" & rs("DatePaid") & "</td>" & chr(13)
            Response.write "            <td>" & rs("WorkDescription") & "</td>" & chr(13)
            Response.write "            <td style=""text-align: center""><a href=""javascript: void(0)"" onclick=""popup('" & URLTC & "ha_timeclockEdit.asp?id=" & rs("id") & iif(Request("id")= "", "", "&empid=" & rs("EmpID")) & "')"" class=""button"">Edit</a>&nbsp;&nbsp;<a class=""button"" href=""javascript:void(0)"" onclick=""window.location.href = 'ha_timeclock.asp?delete=" & rs("id") & "&id=" & eEmpID & "'"">Remove</a></td>" & chr(13)
            Response.write "          </tr>" & chr(13)
            count = count + 1
            rs.MoveNext
          Loop
          rs.close
          Set rs = Nothing

          
        %>
        </table>
        <br /><br />
        <div class="noPrint" style="margin: 0 auto; width: 80%">
        <%
	      ' page navigation
          Response.Write GetHitCountAndPageLinks(Request.ServerVariables("URL"), intRecordCount, recordsPerPage, numOfPages, currentPage, qString)
        %>
        </div>
        <% 
          Else
            Response.write "<div style=""text-align: center; font-weight: bold"">There are no records currently for this employee.</div>"
          End If 
        %>
      </td>
    </tr>
    <tr>
      <td>
        <br />
        <br />
        <br />
        <br />
      </td>
    </tr>
  </table>
    <!--#include virtual="ha_BackOffice/includes/BodyParts/ha_PageTail.inc"-->
</body>
</html>
<%
  Function iif(condition, consequent, alternative)
    If condition Then iif=consequent Else iif=alternative
  End Function

  Function CheckNumber(num)
    If num MOD 3 = 0 Then
      CheckNumber = " style=""background-color: #B4D7BF;"""
    Else
      CheckNumber = ""
    End If
  End Function

  Function GetHitCountAndPageLinks(sHref, nTotalRecords, nRecordsPerPage, intPageCount, currentPage, sQueryString)
	Dim sReturnValue

	sReturnValue = vbCrLf _
		& "<table align=""center"" border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">" & vbCrLf _
		& "<tr>" & vbCrLf _
		& "<td align=""center""><font style=""color: Black""><b>" _
              & DisplayHitCount(nTotalRecords, nRecordsPerPage, intPageCount, currentPage) _
              & "</b></font><br></td>" & vbCrLf _
		& "</tr>" & vbCrLf

	If intPageCount > 1 Then
		sReturnValue = sReturnValue & "<tr>" & vbCrLf _
			& "<td>" & vbCrLf _
			& DisplayPageLinks(sHref, sQueryString, intPageCount, currentPage) & vbCrLf _
			& "</td>" & vbCrLf _
			& "</tr>" & vbCrLf
	End If

	sReturnValue = sReturnValue & "</table>" & vbCrLf

	GetHitCountAndPageLinks = sReturnValue
  End Function


  Function DisplayHitCount(nTotalRecords, nRecordsPerPage, intPageCount, currentPage)
	Dim sReturnValue

	If nTotalRecords = 0 Then
		sReturnValue = "No Results Found."
	Elseif nTotalRecords = 1 then
		sReturnValue = "1 Record Found."
	Else
		Dim nLowestDisplayed, nHighestDisplayed
		Call GetHighAndLowRecordsDisplayed(nTotalRecords, nRecordsPerPage, intPageCount, currentPage, nLowestDisplayed, nHighestDisplayed)

		If nLowestDisplayed = nHighestDisplayed Then
			sReturnValue = nTotalRecords & " Records Found -- Showing Result " & nLowestDisplayed
		Else
			sReturnValue = nTotalRecords & " Records Found -- Showing Results " & nLowestDisplayed & " to " & nHighestDisplayed
		End If
	End If

	DisplayHitCount = sReturnValue
  End Function


  Sub GetHighAndLowRecordsDisplayed(nTotalRecords, nRecordsPerPage, intPageCount, currentPage, nLowestDisplayed, nHighestDisplayed)
	nLowestDisplayed = currentPage * nRecordsPerPage -19

	If currentPage >= intPageCount and currentPage = intPageCount Then		' The last page is displayed
		' The highest record displayed is the last record
		nHighestDisplayed = nTotalRecords
	Else
		nHighestDisplayed = (currentPage) * nRecordsPerPage
	End If
  End Sub

  Function DisplayPageLinks(sHref, sQueryString, intPageCount, currentPage)
	If intPageCount = 1 Then
		DisplayPageLinks = ""
		Exit Function
	End If

	Dim sReturnValue

	sReturnValue = "<table cellpadding=""0"" cellspacing=""1"" style=""width: 100%; border: 0;"">" & VbCrLf _
		& VbCrLf _
		& "<tr valign=""bottom"">" & VbCrLf _
		& "<td width=""25%"" nowrap=""nowrap"">" & VbCrLf

	If currentPage > 1 Then
		sReturnValue = sReturnValue & _
                "<a href=""" & sHref & "?page=" & currentPage - 1 & sQueryString & """ class=""listing"">" & _
                "&nbsp;&nbsp;&lt;&lt; Previous</a>&nbsp;&nbsp;&nbsp;<BR>" & vbCrLf
	Else
		sReturnValue = sReturnValue & "&nbsp;"	
	End If

	sReturnValue = sReturnValue & "</td>" & VbCrLf _
		& "<td align=""center"" valign=""middle"" nowrap=""nowrap"">" & VbCrLf _
		& "<font size=""-1"" color=""#666666"">" & VbCrLf

	Dim nLowPageLink, nHighPageLink, i

	' Make initial high/low calculations
	nLowPageLink = currentPage - 3
	nHighPageLink = currentPage + 3

	' If low is less than zero, increase both values
	If nLowPageLink < 1 Then
		nHighPageLink = nHighPageLink - nLowPageLink
		nLowPageLink = 1
	End If

	' If high is greater than total pages, decrease both values
	If nHighPageLink > intPageCount Then
		nLowPageLink = nLowPageLink - (nHighPageLink - intPageCount)
		nHighPageLink = intPageCount
	End If

	' Ensure that the low value is not negative
	If nLowPageLink < 1 Then nLowPageLink = 1

	If nLowPageLink < currentPage Then
		' Display a link to the first page
		sReturnValue = sReturnValue & "<a href=""" & sHref & "?page=1" & sQueryString & """ class=""listing"">" & _
                "First . . .</a>" & vbCrLf
	End If
			
	For i = nLowPageLink To nHighPageLink
		If i = currentPage Then
			sReturnValue = sReturnValue & "<B>&nbsp;" & i & "&nbsp;</B>" & VbCrLf
		Else
			sReturnValue = sReturnValue & _
                  "&nbsp;<a href=""" & sHref & "?page=" & i & sQueryString & """ class=""listing"">" & _
                  i & "</a>&nbsp;" & vbCrLf
		End If
	Next

	If currentPage < intPageCount Then
		' Display a link to the last page
		sReturnValue = sReturnValue & _
            "<a href='" & sHref & "?page=" & intPageCount & sQueryString &  "' class=""listing"">" & _
            ". . . Last</a> " & vbCrLf
	End If

	sReturnValue = sReturnValue & "</FONT>" & VbCrLf _
		& "</td>" & VbCrLf _
		& "<td align=""right"" width=""25%"" nowrap=""nowrap"">" & VbCrLf

	If currentPage < intPageCount Then
		sReturnValue = sReturnValue & _
            "&nbsp;&nbsp;&nbsp;" & _
            "<a href=""" & sHref & "?page=" & currentPage + 1 & sQueryString & """ class=""listing"">" & _
            "Next &gt;&gt;&nbsp;&nbsp;</a><br>" & vbCrLf
	Else
		sReturnValue = sReturnValue & "&nbsp;"	
	End If

	sReturnValue = sReturnValue & "</td>" & VbCrLf _
		& "</tr>" & VbCrLf _
		& VbCrLf _
		& "</table>" & VbCrLf

	DisplayPageLinks = sReturnValue
  End Function
%>