<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<%
    ThisPage = "https://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL")
    
    dim weSunday
    dim pdWednesday
    
    weSunday = (dateadd("d", (1 - datepart("w", date())), date()))
    pdWednesday = dateadd ("d", 3, weSunday)
        
    'dim todaysLongDate
    dim eWeekEnding
    dim ePayDate
        
    'todaysLongDate=NOW()
    'todaysJustDate=FormatDateTime(todaysLongDate,2)
        
    if Request("WeekEnding") <> "" then
        eWeekEnding = Request("WeekEnding")
                ePaydate = dateadd ("d", 3, eWeekEnding)

    else
        eWeekEnding = weSunday
        ePayDate = pdWednesday
   end if
%>

<!--#include virtual="includes/connection.asp"-->

<%
  If session("LoggedIn")<>"True" OR isNULL(session("LoggedIn")) THEN
    Response.Redirect "ha_login.asp"
  End If
%>

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>Harvest American, Inc. | Payroll Administration - Summary</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <link href="<%=strPathMenuIncludes%>ha_menu.css" type="text/css" rel="stylesheet">
    <link href="<%=strPathMenuIncludes%>/calendar.css" rel="stylesheet" type="text/css">
    
    <script type="text/javascript" src="<%=strPathMenuIncludes%>/js/calendar_date_picker.js"></script>

    <script type="text/javascript">
        function confirmAction() {
            return confirm("Do you really want to mark all paid?");
        }

        //Calendar Date Picker
        var cdp1 = new CalendarDatePicker();
        var props = {
            debug: true,
            excludeDays: [6, 1, 2, 3, 4, 5],
            formatDate: '%m/%d/%y'
        };
        var cdp2 = new CalendarDatePicker(props);
        props.formatDate = '%m/%d/%y';
        var cdp3 = new CalendarDatePicker(props);
        cdp3.endYear = cdp3.startYear + 1;
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
                        SQL="ha_PayrollSummary('" & eWeekEnding &"')"

                        rs.open SQL,objcon,2,3

            %>

            <body>

                <table width="450" border="0" cellspacing="1" cellpadding="0" align="center">
                    <tr>
                        <td colspan="2">
                            <form method="get" action="ha_payrollsummary.asp" name="frm">
                            <table style="margin: 5px 0 0 -9px;" width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
                                <tr>
                                    <td align="right" width="40%" valign="baseline">
                                        <h4>WEEK ENDING:&nbsp;</h4>
                                    </td>
                                    <td width="18%" valign="baseline" align="center">
                                        <input type="Text" style='text-align: right; width: 5em;' name="WeekEnding" value="<%=eWeekEnding%>" size="7" readonly>
                                    </td>
                                    <td width="10" align="left">
                                        <a href="#" onclick="cdp2.showCalendar(this, 'WeekEnding'); return false;" title="View Calendar">
                                            <img style="padding-left: 5px;" align="absmiddle" src="<%=strPathMenuImages%>/calendar.png" border="0"></a>
                                    </td>
                                    <td align="left" width="32">
                                        <input type="submit" name="sub" value="Refresh" class="button">
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
                                        <h4>There are no unpaid records for the week ending<br>
                                            <%=FormatDateTime(eWeekEnding, 1)%></h4>
                                    </td>
                                </tr>
                            </table>
                            <%
                                                else
                                                do while not rs.eof
                            %>

                            <table class="outlined" style="border-color: #A1C94F; margin: 10px 0 0 -9px;" width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
                                <tr>
                                    <td>

                                        <table width="450" height="30" border="0" cellspacing="0" cellpadding="1" align="center">
                                            <tr bgcolor="#A1C94F">
                                                <th width="20%" align='center'>EmpID <% If trim(Session("ha_login")) = "lrowlands" Or trim(Session("ha_login")) = "Hans"Then %>
                                                    <a href="../timeclock/ha_timeclock.asp?EmpID=<%=rs("EmpID")%>"><%=rs("EmpID")%></a>
                                                    <% Else %>
                                                    <%=rs("EmpID")%>
                                                    <% End If %>
                                                </th>
                                                <th align="left" style="padding-left: 25px;">
                                                    <%=rs("FirstName")%>&nbsp;<%=rs("LastName")%>
                                                </th>
                                                <th width="25%" align='center'>
                                                    <span style='color; #336600; font-size: 10px; font-weight: bold;'>Total: </span><%=rs("SumTotal")%>
                                                </th>
                                            </tr>
                                        </table>

                                        <table width="450" border="0" cellspacing="0" cellpadding="0" align="center">
                                            <tr bgcolor="   #CDE2A1" height="22">
                                                <td width="20%" align='center'>
                                                    <span style='color; #336600; font-size: 11px;'>Regular</span>
                                                </td>
                                                <td width="20%" align="center">
                                                    <span style='color; #336600; font-size: 11px;'>Holiday</span>
                                                </td>
                                                <td width="20%" align="center">
                                                    <span style='color; #336600; font-size: 11px;'>Vacation</span>
                                                </td>
                                                <td width="20%" align='center'>
                                                    <span style='color; #336600; font-size: 11px;'>Sick Time</span>
                                                </td>
                                                <td width="20%" align='center'>
                                                    <span style='color; #336600; font-size: 11px;'>Mileage</span>
                                                </td>
                                            </tr>
                                        </table>
                                        <table width="450" border="0" cellspacing="0" cellpadding="0" align="center">
                                            <tr height="22">
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
                        <h2>Payroll<br>
                            Action Menu</h2>
                    </td>
                </tr>
                <tr>
                    <td width="18%" valign="baseline" align="center">
                        <input type="Text" style='text-align: right;' name='PayDate' value="<%=ePayDate%>" size="7" readonly>
                    </td>
                </tr>
                <tr>
                    <td align="center" width="32">
                        <input type="submit" name="sub" value="Mark All Paid" class="button">
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

    <table id="calendarTable">
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
    </table>
    <!-- End of the Calendar //-->

</body>
</html>
