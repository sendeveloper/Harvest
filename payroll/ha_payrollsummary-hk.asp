<!DOCTYPE html>
<%
    ThisPage = "https://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL")
    
    dim weSunday, pdWednesday, fmtDate
    
    weSunday = (dateadd("d", datepart("w", date()), date()) - Weekday(dateadd("d", datepart("w", date()), date()), vbSunday))
    pdWednesday = dateadd ("d", 3, weSunday)
    fmtDate = Split(weSunday,"/")
        
    'dim todaysLongDate
    dim eWeekEnding
    dim ePayDate
        
    'todaysLongDate=NOW()
    'todaysJustDate=FormatDateTime(todaysLongDate,2)
    
    If Request("date") = "" then
      eWeekEnding = fmtDate(2) & "-"
      If fmtDate(0) < 10 Then
        eWeekEnding = eWeekEnding & "0" & fmtDate(0) & "-"
      Else
        eWeekEnding = eWeekEnding & fmtDate(0) & "-"
      End If

      If fmtDate(1) < 10 Then
        eWeekEnding = eWeekEnding & "0" & fmtDate(1)
      Else
        eWeekEnding = eWeekEnding & fmtDate(1)
      End If
      ePayDate = pdWednesday
    Else
      eWeekEnding = Request("date")
      ePaydate = dateadd ("d", 3, eWeekEnding)   
    End if

    If Not Request("type") = "" Then
      Session("type") = Request("type")
    Else
      If Session("type") = "" Then Session("type") = "unpaid"
    End If
%>

<!--#include virtual="includes/connection.asp"-->

<%
  If session("LoggedIn")<>"True" OR isNULL(session("LoggedIn")) THEN
    Response.Redirect "ha_login.asp"
  End If
%>

<html>
<head>
  <title>Harvest American, Inc. | Payroll Administration - Summary</title>
  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
  <link href="<%=strPathMenuIncludes%>ha_menu.css" type="text/css" rel="stylesheet" />
  <link href="http://code.jquery.com/ui/1.10.2/themes/south-street/jquery-ui.css" rel="stylesheet" />
<!--  <script type="text/javascript" src="<%=strPathIncludes%>jq-cal/js/scripts.js"></script>-->
  <!--<link href="<%=strPathMenuIncludes%>/calendar.css" rel="stylesheet" type="text/css" />-->
  <!--#include virtual="menu/includes/menu.inc"-->
  <!--<script type="text/javascript" src="<%=strPathMenuIncludes%>/js/calendar_date_picker.js"></script>-->

  <script type="text/javascript">
    function confirmAction() {
      return confirm("Do you really want to mark all paid?");
    }

    //Calendar Date Picker
    //var cdp1 = new CalendarDatePicker();
    //var props = {
    //  debug: true,
    //  excludeDays: [6, 1, 2, 3, 4, 5],
    //  formatDate: '%m/%d/%y'
    //};
    //var cdp2 = new CalendarDatePicker(props);
    //props.formatDate = '%m/%d/%y';
    //var cdp3 = new CalendarDatePicker(props);
    //cdp3.endYear = cdp3.startYear + 1;
  </script>

</head>
<body>

  <!--#include virtual="menu/includes/ha_header.inc"-->

  <!--#include virtual="menu/includes/ha_menu.inc"-->

  <div id="content">
    <div id="columnA">
      <%

        Dim rs
        'Dim SQL
                                                
        set rs = server.createObject("ADODB.Recordset")
        SQL="ha_PayrollSummary_sorttype('" & eWeekEnding &"','" & Session("type") & "')"
        'response.write SQL
        rs.open SQL,objcon,2,3

      %>

      <table width="450" border="0" cellspacing="1" cellpadding="0" align="center">
        <tr>
          <td colspan="2">
            <form method="get" action="ha_payrollsummary-hk.asp" name="frm">
              <table style="margin: 5px 0 0 -9px;" width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
                <tr>
                  <td align="right" width="40%" valign="baseline">
                    <h4>WEEK ENDING:&nbsp;</h4>
                  </td>
                  <td width="18%" valign="baseline" align="center">
                    <!--<input type="text" style='text-align: right; width: 5em;' id="weekending" name="WeekEnding" value="<%=eWeekEnding%>" size="7" readonly="readonly" />-->
                    <input type="date" name="date" id="date" class="datepick" value="<%=eWeekEnding%>" />
                    <script type="text/javascript" src="http://code.jquery.com/jquery-1.9.1.js"></script>
                    <script type="text/javascript" src="http://code.jquery.com/ui/1.10.2/jquery-ui.js"></script>
                    <script type="text/javascript">
                      (function () {
                          var elem = document.createElement('input');
                          elem.setAttribute('type', 'date');

                          if (elem.type === 'text') {
                              $('.datepick').datepicker({
                                  dateFormat: 'yy-mm-dd'
                              });
                          }
                      })(); 
                    </script>
                  </td>
                  <td width="10" align="left">
                    <!--<a href="#" onclick="cdp2.showCalendar(this, 'date'); return false;" title="View Calendar">
                    <img alt="" style="padding-left: 5px;" align="absmiddle" src="<%=strPathMenuImages%>/calendar.png" border="0" /></a>-->&nbsp;<input type="submit" name="sub" value="Refresh" class="button" />
                  </td>
                </tr>
                <tr>
                  <td colspan="2" style="text-align: right; vertical-align: top">Payroll Type:&nbsp;</td>
                  <td style="height: 3em; width: 100%; float: right">
                                    
                  <select id="selPayrollType" style="width: 9.25em; margin: 0 auto" onchange="window.location='<%request.servervariables("URL")%>?type=' + this.value + '&WeekEnding=' + weekending.value ">
                    <option value="paid"<%=iif(Session("type")="paid"," selected='selected'","")%>>Paid Records</option>
                    <option value="NULL"<%=iif(Session("type")="NULL"," selected='selected'","")%>>Unpaid Records</option>
                    <option value="both"<%=iif(Session("type")="both"," selected='selected'","")%>>All Records</option>
                  </select>
                </td>
              </tr>
            </table>
          </form>

          <%
            if rs.eof then
          %>
          <table style="margin: 10px 0 0 -9px;" width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
            <tr>
              <td align='center'>
                <h4>There are no unpaid records for the week ending<br />
                <%=FormatDateTime(eWeekEnding, 1)%></h4>
              </td>
            </tr>
          </table>
          <%
          else
          %>
          <div style="clear: both">&nbsp;</div>
          <%
          
          do while not rs.eof
          %>
          <table class="outlined" style="border-color: #A1C94F; margin: 10px 0 0 -9px;" width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
            <tr>
              <td>
                <table style="width: 450px; height: 30px; border: 0; border-spacing: 0; padding: 1px" align="center">
                  <tr bgcolor="#A1C94F">
                    <th width="20%" align='center'>EmpID <% If trim(Session("ha_login")) = "lrowlands" Or trim(Session("ha_login")) = "Hans"Then %>
                      <a href="../timeclock/ha_timeclock.asp?EmpID=<%=rs("EmpID")%>"><%=rs("EmpID")%></a>
                      <% Else %>
                      <%=rs("EmpID")%>
                      <% End If %>
                    </th>
                    <th align="left" style="padding-left: 25px;">
                    <%=rs("LastName")%>,&nbsp;<%=rs("FirstName")%>
                    </th>
                    <th width="25%" align='center'>
                      <span style='color: #336600; font-size: 10px; font-weight: bold;'>Total: </span><%=rs("SumTotal")%>
                    </th>
                  </tr>
                </table>

                <table width="450" border="0" cellspacing="0" cellpadding="0" align="center">
                  <tr bgcolor="#CDE2A1" style="height: 22px">
                    <td width="20%" align='center'>
                      <span style='color: #336600; font-size: 11px;'>Regular</span>
                    </td>
                    <td width="20%" align="center">
                      <span style='color: #336600; font-size: 11px;'>Holiday</span>
                    </td>
                    <td width="20%" align="center">
                      <span style='color: #336600; font-size: 11px;'>Vacation</span>
                    </td>
                    <td width="20%" align='center'>
                      <span style='color: #336600; font-size: 11px;'>Sick Time</span>
                    </td>
                    <td width="20%" align='center'>
                      <span style='color: #336600; font-size: 11px;'>Mileage</span>
                    </td>
                  </tr>
                </table>
                <table width="450" border="0" cellspacing="0" cellpadding="0" align="center">
                  <tr style="height:22px">
                    <td width="20%" align='center'>
                      <%=rs("SumRegular")%>
                    </td>
                    <td width="20%" align="center">
                      <%=rs("SumHoliday")%>
                    </td>
                    <td width="20%" align="center">
                      <%=rs("SumVacation")%>
                    </td>
                    <td width="20%" align='center'>
                      <%=rs("SumPaidTimeOff")%>
                    </td>
                    <td width="20%" align='center'>
                      <%=rs("Mileage")%>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>
          </table>

          <%
          rs.MoveNext
          loop
          end if
          %>

          </td>
        </tr>
      </table>
    </div>

    <div id="columnB">
      <form method="post" action="ha_payrolltimeclockPost.asp" name="frm2" onsubmit="return confirmAction()">
        <table width="180" border="0" cellspacing="1" cellpadding="1" align="center">
          <tr>
            <td align="center">
              <h2>Payroll<br />
              Action Menu</h2>
            </td>
          </tr>
          <tr>
            <td width="18%" valign="baseline" align="center">
              <input type="text" style='text-align: right;' name='PayDate' value="<%=ePayDate%>" size="7" readonly="readonly" />
            </td>
          </tr>
          <tr>
            <td align="center" width="32">
              <input type="submit" name="sub" value="Mark All Paid" class="button" />
            </td>
          </tr>
        </table>
      </form>

      <!--#include virtual="menu/includes/ha_domains.inc"-->

    </div>
    <div style="clear: both;">&nbsp;</div>
  </div>

  <!--#include virtual="menu/includes/ha_footer.inc"-->

  <!--
        Start of the Calendar

        Script written by Martijn Korse

        http://devshed.excudo.net
//-->

  <!--<table id="calendarTable">
    <tbody id="calendarTableHead">
      <tr>
        <td colspan="4" align="left">
          <select id="selectMonth">
            <option value="0">January</option>
            <option value="1">February</option>
            <option value="2">March</option>
            <option value="3">April</option>
            <option value="4">May</option>
            <option value="5">June</option>
            <option value="6">July</option>
            <option value="7">August</option>
            <option value="8">September</option>
            <option value="9">October</option>
            <option value="10">November</option>
            <option value="11">December</option>
          </select>
        </td>
        <td colspan="2" align="center">
          <select id="selectYear"></select></td>
        <td align="right"><a href="#" id="closeCalendarLink">X</a></td>
      </tr>
    </tbody>
    <tbody id="calendarTableDays">
      <tr id="calenderDaysIndex">
        <td>Su</td>
        <td>Mo</td>
        <td>Tu</td>
        <td>We</td>
        <td>Thu</td>
        <td>Fr</td>
        <td>Sa</td>
      </tr>
    </tbody>
    <tbody id="calendar"></tbody>
  </table>-->
  <!-- End of the Calendar //-->

</body>
</html>
<%
Function iif(condition, consequent, alternative)
    If condition Then iif=consequent Else iif=alternative
End Function
%>