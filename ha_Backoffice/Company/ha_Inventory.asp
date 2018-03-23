<!DOCTYPE html>
<!--#include virtual="includes/connection.asp"-->
<!--#include virtual="ha_BackOffice/includes/Config.asp"-->
<%
  Dim connString, recordsPerPage, currentPage, strSQL, numOfPages, OrderBy, AscDesc  
  ColorTab = 5
  PageHeading = "IT Inventory"

%>
<html>
<head>
  <!--include file="sql.asp"-->
  <script type="text/javascript" src="haFunctions.js"></script>
  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
  <link href="calendar.css" rel="stylesheet" type="text/css" />
  <script type="text/javascript" src="calendar_date_picker.js"></script>
  <title>Harvest American Backoffice - Employees</title>
  <link href="<%=strPathIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet" media="screen" />
  <link href="ha_Print.css" type="text/css" rel="stylesheet" media="print" />
  <link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>menu/dropdowntabfiles/halfmoontabs.css" media="screen" />
  <script type="text/javascript" src="<%=strPathIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>
  <style type="text/css">
    .button
      {
        font-weight: bold;
        font-size: 1em;
        font-family: Verdana, Arial, Helvetica, sans-serif;
        color: #336600;
        padding: .25em .5em;
        background-color: #CDE2A1;
        border-top: .125em solid #C0C0C0;
        border-right: .125em solid black;
        border-bottom: .125em solid black;
        border-left: .125em solid #C0C0C0;
        text-align: center;
        text-decoration: none;
        margin-bottom: 2em;
      }
    
    .button:hover
      {
        font-weight: bold;
        font-size: 1em;
        font-family: Verdana, Arial, Helvetica, sans-serif;
        color: white;
        background-color: #336600;
        border-color: black #C0C0C0 #C0C0C0 black;
      }

    .textbox
      {
        display: inline-block;
        height: 1.5em;
        width: 13.3em;
        margin-bottom: .5em;
      }

    .checkbox
      {
        margin-bottom: .5em;
      }

    select
      {
        display: inline-block;
        height: 1.5em;
        margin-bottom: .5em;
      }
  </style>
</head>
<body class="gray_desktop">
  <form id="frm" action="ha_Products.asp" method="post">
    <table class="MainBody" style="margin: 20px auto 0; width: 1200px !important; padding: 0; border-spacing: 0;">
      <tr>
        <td>
          <div class="noPrint">
            <!--#include virtual="ha_Backoffice/includes/bodyparts/ha_PageHeading.inc"-->
          </div>
        </td>
      </tr>
      <tr>
        <td>
          <%
            Dim rs, sqlText, sqlSort, sqlPageSize, intRecordCount, count: count = 1
    
            set rs = server.createObject("ADODB.Recordset")
          %>
          <div>
            <div style="clear: both; margin: 2em 0"></div>
              
                <%

	                ' set number of records per page
	                recordsPerPage = 20

	                ' check for page
	                If Request.QueryString("Page") = "" Then
		              currentPage = 1
	                Else
		              currentPage = CINT(Request.QueryString("Page")) ' convert Page request to an integer
	                End If

                    ' SQL Query
                    strSQL = "SELECT * FROM ha_Inventory"

	                rs.Open strSQL, connCasper10, 1, 3

              If Not rs.EOF Then
              %>
              <table style="width: 80%; margin: 0 auto" cellpadding="5" cellspacing="5">
                <tr style="background-color: #008000">
                  <th style="width: 10%; color: White">ID</th>
                  <th style="width: 20%; color: White">Description</th>
                  <th style="width: 15%; color: White">Assigned To</th>
                  <th style="width: 15%; color: White">Location</th>
                  <th style="width: 25%; color: White">Notes</th>
                  <th style="width: 15%; color: White"><span class="noprint">Options</span></th>
                </tr>
                <%
	              rs.PageSize = recordsPerPage

	              numOfPages = rs.PageCount
                  intRecordCount = rs.RecordCount
	
	              ' set current page
	              rs.AbsolutePage = currentPage
	
	                ' loop through the table to obtain records
                  Do While Not count = 21 And Not rs.EOF
                %>
                <tr<%=CheckNumber(count)%>>
                  <td><%=rs("ID")%></td>
                  <td><%=rs("Description")%></td>
                  <td><%=rs("AssignedTo")%></td>
                  <td><%=rs("Location")%></td>
                  <td><%=rs("Notes")%></td>
                  <td style="text-align: center">
                    <span class="noPrint"><a href="javascript:void(0)" style="text-decoration: none" onclick="openPopUp('ha_invitemview.asp?id=<% =rs("ID") %>&itype=view')">Edit</a></span>
                  </td>
                </tr>
                <%
                    count = count + 1
	                ' move to the next record
	                rs.MoveNext
	              Loop
	
	              ' close the recordset and connection object then clean out
	              rs.close	                
                %>
              </table>
              <br />
              <div class="noPrint" style="width: 80%; margin: 0 auto; font-size: 1.5em">
                <%
	              ' page navigation
                  Response.Write GetHitCountAndPageLinks(Request.ServerVariables("URL"), intRecordCount, recordsPerPage, numOfPages, currentPage, "")
                %>
              </div>
              <% End If %>
              <br /><br />

              <a class="noPrint button" href="javascript: void(0)" onclick="window.print()" style="margin:  0 0 0 50em">Print</a>&nbsp;&nbsp;&nbsp;<a href="javascript:void(0);" onclick="openPopUp('ha_invitemview.asp?iMode=Add')" class="noPrint button" style="text-align: center">Add Item</a>
            <br />
          </div>
        </td>
      </tr>
      <tr>
        <td style="height: 2em; vertical-align: middle;">
          
        </td>
      </tr>
      <tr>
        <td>&nbsp;</td>
      </tr>
    </table>
    <!--#include virtual="ha_BackOffice/includes/BodyParts/ha_PageTail.inc"-->
  </form>
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
		& "<table align=""center"" border=""0"" cellpadding=""0 cellspacing=""0"" width=""80%"">" & vbCrLf _
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