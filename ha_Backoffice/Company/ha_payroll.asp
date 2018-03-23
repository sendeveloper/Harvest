<!DOCTYPE html>
<html>
<head>
  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/Config.asp"-->
  <!--#include file="sql.asp"-->
  <!--#include file="lib.asp"-->
  <%
    dim weSunday, pdWednesday, eWeekEnding, ePayDate, ex, numAdd, dob
    ex = WeekdayName(Weekday(date(),1))
    'response.write ex
    
    Select Case ex
      Case "Monday"
      numAdd = 6
      Case "Tuesday"
      numAdd = 5
      Case "Wednesday"
      numAdd = 4
      Case "Thursday"
      numAdd = 3
      Case "Friday"
      numAdd = 2
      Case "Saturday"
      numAdd = 1
      Case "Sunday"
      numAdd = 0
    End Select

    weSunday = dateadd("d", numAdd, date())
    pdWednesday = dateadd("d", 3, weSunday)
        
    'todaysLongDate=NOW()
    'todaysJustDate=FormatDateTime(todaysLongDate,2)
    
    If Request("weekending") = "" then
      eWeekEnding = FormatDate(weSunday)
      ePayDate = FormatDate(pdWednesday)
    Else
      eWeekEnding = Request("weekending")
      'eWeekEnding = ex(0)
      ePaydate = FormatDate(dateadd("d", 3, eWeekEnding))
    End if

    Function FormatDate(aDate)
      Dim fmtDate, newDate: fmtDate = Split(aDate,"/")

      newDate = fmtDate(2) & "-"
      If fmtDate(0) < 10 Then
        newDate = newDate & "0" & fmtDate(0) & "-"
      Else
        newDate = newDate & fmtDate(0) & "-"
      End If

      If fmtDate(1) < 10 Then
        newDate = newDate & "0" & fmtDate(1)
      Else
        newDate = newDate & fmtDate(1)
      End If

      FormatDate = newDate
    End Function

    If Not Request("type") = "" Then
      Session("type") = Request("type")
    Else
      Session("type") = 0
    End If

    Session("Redirect") = ""
	ColorTab = 5
    PageHeading = "Payroll"

    Function CheckNumber(num)
      If num MOD 3 = 0 Then
        CheckNumber = " style=""background-color: #B4D7BF"""
      End If
    End Function
  %>
  <title>Harvest American - <%=PageHeading%></title>
  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
  <link href="<%=strPathIncludes %>datepick/jquery-ui.css" rel="stylesheet" />
  <link href="<%=strPathIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet" />
  <link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>menu/dropdowntabfiles/halfmoontabs.css" />
  <script type="text/javascript" src="<%=strPathIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>
  <style type="text/css">
    tr.calrowhead { background-color: #008000; }
  </style>
</head>
<body class="gray_desktop">
  <table cellpadding="0" cellspacing="0" class="MainBody" style="margin: 16px auto 0; width: 1200px">
    <tr>
      <td>
        <!--#include virtual="ha_Backoffice/includes/bodyparts/ha_PageHeading.inc"-->
      </td>
    </tr>
    <tr>
      <td>
        <table style="width: 450px; margin: 0 auto" cellpadding="5">
          <tr class="calrowhead">
            <th style="color: white;" colspan="2">Payroll Action Menu</th>
          </tr>
          <tr>
            <td style="text-align: right">Date to Mark as Paid: </td>
            <td>
              <form method="post" action="http://www.harvestamerican.info/payroll/ha_payrolltimeclockPost.asp" name="frm2" onsubmit="return confirmAction()">
                <input type="date" name="PayDate" id="PayDate" class="datepick" value="<%=ePayDate%>" />
                &nbsp;&nbsp;&nbsp;&nbsp;
                <a href="javascript:void(0)" onclick="document.frm2.submit()" class="button">Mark All Paid</a>
              </form>
            </td>
          </tr>
          <tr>
            <td style="text-align: right">Week Ending: </td>
            <td>
              <form method="get" action="" name="frm">
                <input type="date" name="weekending" id="weekending" class="datepick" value="<%=eWeekEnding%>" />
                &nbsp;&nbsp;&nbsp;&nbsp;
                <a href="javascript:void(0)" onclick="document.frm.submit();" class="button"> Refresh </a>
                <script type="text/javascript" src="http://code.jquery.com/jquery-1.9.1.js"></script>
                <script type="text/javascript" src="http://code.jquery.com/ui/1.10.2/jquery-ui.js"></script>
                <script type="text/javascript">
                (function () {
                  var elem = document.createElement('input');
                  elem.setAttribute('type', 'date');

                  if (elem.type === 'text') {
                    $('.datepick').datepicker({
                    dateFormat: 'yy-mm-dd',
                    showOn: 'button',
                    buttonImage: '../images/cal.gif',
                    buttonImageOnly: true
                    });
                  }
                })(); 
                </script>
              </form>
            </td>
          </tr>
          <tr>
            <td style="text-align: right">Payroll Type: </td>
            <td>
              <select id="selPayrollType" style="width: 148px; margin: 0 auto" onchange="window.location='<%request.servervariables("URL")%>?type=' + this.value + '&weekending=' + weekending.value ">
                <option value="1"<%=iif(Session("type")="1"," selected=""selected""","")%>>Paid Records</option>
                <option value="0"<%=iif(Session("type")="0"," selected=""selected""",iif(Request("type")=""," selected=""selected""",""))%>>Unpaid Records</option>
                <option value="2"<%=iif(Session("type")="2"," selected=""selected""","")%>>All Records</option>
              </select>
            </td>
          </tr>
        </table>
        <br /><br />
        <%
          Dim rs, strsql, count: count = 1                                             
          strsql = "EXEC ha_Payroll_hk '" & eWeekEnding &"', " & Session("type")
          'response.write strsql
          Set rs = Server.CreateObject("ADODB.Recordset")
          rs.Open strsql, connCasper10, 1, 3
          'response.write strsql
          If Not rs.EOF Then
          %>
          <table style="width: 960px;" align="center" cellpadding="5" cellspacing="5">
            <tr class="calrowhead">
              <th rowspan="2" style="color: white;">Emp ID</th>
              <th rowspan="2" style="color: white; width: 300px">Name</th>
              <th rowspan="2" style="color: white;">Birthday</th>
              <th style="color: white;" colspan="5">Type of Pay</th>
            </tr>
            <tr class="calrowhead">
              <th style="color: white;">Regular</th>
              <th style="color: white;">Vacation</th>
              <th style="color: white;">Holiday</th>
              <th style="color: white;">Personal</th>
              <th style="color: white;">Total</th>
            </tr>

          <%
            Do While Not rs.EOF
              Response.write "            <tr" & CheckNumber(count) & ">" & chr(13)
              Response.write "              <td style=""text-align: center"">" & chr(13)
              Response.write "                " & rs("EmpID") & chr(13)
              Response.write "              </td>" & chr(13)
              Response.write "              <td>" & chr(13)
              Response.write "                <a href=""ha_timeclock.asp?id=" & rs("EmpID") & """>" & rs("LastName") & ",&nbsp;" & rs("FirstName") & "</a>" & chr(13)
              Response.write "              </td>" & chr(13)
              Response.write "              <td style=""text-align: center"">" & chr(13)
              If Not rs("DOB") = "1900-01-01 00:00:00" Then
              Response.write "                " & Month(rs("DOB")) & "/" & Day(rs("DOB")) & chr(13)
              End If
              Response.write "              </td>" & chr(13)
              Response.write "              <td style=""text-align: center"">" & chr(13)
              Response.write "                " & rs("Regular") & chr(13)
              Response.write "              </td>" & chr(13)
              Response.write "              <td style=""text-align: center"">" & chr(13)
              Response.write "                " & rs("Holiday") & chr(13)
              Response.write "              </td>" & chr(13)
              Response.write "              <td style=""text-align: center"">" & chr(13)
              Response.write "                " & rs("Vacation") & chr(13)
              Response.write "              </td>" & chr(13)
              Response.write "              <td style=""text-align: center"">" & chr(13)
              Response.write "                " & rs("Personal") & chr(13)
              Response.write "              </td>" & chr(13)
              Response.write "              <td style=""text-align: center"">" & chr(13)
              Response.write "                " & rs("TotalHrs") & chr(13)
              Response.write "              </td>" & chr(13)
              Response.write "            </tr>" & chr(13)
              count = count + 1
              rs.MoveNext
            Loop
          %>
          </table>
          <%
          Else
          %>
            <span style="width: 400px; text-align: center; margin: 0 auto">There are no records to be found for the Week Ending <%=eWeekEnding %>.</span>
          <%
          End If
        %>
      </td>
    </tr>
    <tr><td><br /><br /></td></tr>
  </table>
</body>
</html>
