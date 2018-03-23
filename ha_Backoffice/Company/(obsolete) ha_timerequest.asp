<!DOCTYPE html>
<html>
<head>
    <%
	Session("Redirect") = ""
	ColorTab = 5
	PageHeading = "Time Off Requests"

    Function CheckNumber(num)
      If num MOD 2 = 0 Then
        CheckNumber = " style=""background-color: #B4D7BF"""
      End If
    End Function


    %>
    <!--#include virtual="includes/connection.asp"-->
    <!--#include virtual="ha_backoffice/includes/Config.asp"-->
    <!--#include file="sql.asp"-->
    <!--#include file="lib.asp"-->
    <title>Harvest American - <%=PageHeading%></title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <link href="<%=strPathIncludes %>datepick/jquery-ui.css" rel="stylesheet" />
    <link href="<%=strPathIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet" />
    <link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>menu/dropdowntabfiles/halfmoontabs.css" />
    <script type="text/javascript" src="<%=strPathIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>
    <style type="text/css" media="print">
        .noPrint
        {
            display: none;
        }
    </style>
    <script type="text/javascript">
        function clickView(id) {
            var URL = 'ha_viewrequest.asp?rMode=View&id=' + id;
            window.open(URL, '_blank', 'width=425; height=500');
        }
        function clickEdit(id) {
            var URL = 'ha_viewrequest.asp?rMode=Edit&id=' + id;
            window.open(URL, '_blank', 'width=425; height=500');
        }
        function clickAdd() {
            var URL = 'ha_viewrequest.asp?rMode=Add';
            window.open(URL, '_blank', 'width=425; height=500');
        }
        function clickChange(date) {
            window.location = 'ha_timerequest.asp?date=' + date;
        }
    </script>
    <style type="text/css">
        tr.calrowhead
        {
            background-color: #008000;
        }
        th.calhead
        {
            width: 14.3%;
            color: white;
            border: 5px solid #E0EEE0;
        }
        
        td.calcell
        {
            border: 5px solid #E0EEE0;
            text-align: right;
            vertical-align: top;
        }
        
        h1
        {
            width: 30%;
            margin-left: 16.5em;
            margin-right: 16.5em;
        }
    </style>
</head>
<body class="gray_desktop">
  <table cellpadding="0" cellspacing="0" class="MainBody" style="margin: 1em auto 0; width: 1200px">
    <tr class="noPrint">
      <td>
        <!--#include virtual="ha_Backoffice/includes/bodyparts/ha_PageHeading.inc"-->
      </td>
    </tr>
    <tr>
      <td>
        <table style="width: 80%; margin: 0 auto;" cellpadding="5">
          <%
            ' declare variables
            Dim yearNum, datePH, monthNum, prevMon, nextMon, assembleDate, dayNum: dayNum = 1

            If Request("date") = "" Then
              datePH = Month(Now) & "/1/" & Year(Now)
              monthNum = Month(datePH)
              yearNum = Year(datePH)
            Else
              datePH = CDate(DateFormat(Month(Request("date")) & "/1/" & Year(Request("date"))))
              monthNum = Month(datePH)
              yearNum = Year(datePH)
            End If

            If monthNum = 1 Then
              prevMon = 12
            Else
              prevMon = monthNum - 1
            End If
            If monthNum = 12 Then
              nextMon = 1
            Else
              nextMon = monthNum + 1
            End If
          %>
          <tr>
            <th colspan="7">
              <a class="button noPrint" href="ha_timerequest.asp?date=<%=monthNum%>-1-<%=yearNum - 1%>"><</a>&nbsp;&nbsp; <a class="button noPrint" href="ha_timerequest.asp?date=<%=prevMon%>-1-<%=iif(monthNum = 1, yearNum - 1, yearNum)%>"><</a> &nbsp;&nbsp;<input type="date" name="changedate" id="changedate" class="datepick noPrint" value="<%=datePH%>" />&nbsp;&nbsp;<a class="button noPrint" href="javascript:void(0)" onclick="clickChange(changedate.value)">Change Date</a> &nbsp;&nbsp; <a class="button noPrint" href="ha_timerequest.asp?date=<%=nextMon%>-1-<%=iif(monthNum = 12, yearNum + 1, yearNum)%>">></a>&nbsp;&nbsp; <a class="button noPrint" href="ha_timerequest.asp?date=<%=monthNum%>-1-<%=yearNum + 1%>">>></a><h1><%=MonthName(monthNum) & " " & yearNum %></h1>
            </th>
          </tr>
          <tr class="calrowhead">
            <th class="calhead">Sunday</th>
            <th class="calhead">Monday</th>
            <th class="calhead">Tuesday</th>
            <th class="calhead">Wednesday</th>
            <th class="calhead">Thursday</th>
            <th class="calhead">Friday</th>
            <th class="calhead">Saturday</th>
          </tr>
          <tr>
            <%
              Dim rs, doW, numberofDays, weekcount, range, count: range = 0: weekcount = 0: count = 1

              doW = Weekday(datePH) - 1

              Do While Not range = doW
            %>
            <td class="calcell">&nbsp;</td>
            <%
                range = range + 1
              Loop

              weekcount = range
              range = 1
              If monthNum = 4 Or monthNum = 6 Or monthNum = 9 Or monthNum = 11 Then
                numberofDays = 30
              ElseIf monthNum = 2 Then
                If monthNum Mod 4 = 0 Then
                  numberofDays = 29
                Else
                  numberofDays = 28
                End If
              Else
                numberofDays = 31
              End If
              
              Do While Not range = numberofDays +1
            %>
            <td class="calcell"><b><%=range %></b><br />
            <%
              datePH = CDate(monthNum & "-" & dayNum & "-" & yearNum)
              Set rs = Server.CreateObject("ADODB.Recordset")
              rs.Open "SELECT * FROM ha_EmpTimeRequest WHERE ([RequestFrom] = '" & datePH & "' OR ([RequestFrom] <= '" & datePH & "' AND [RequestTo] >= '" & datePH & "')) AND [IsApproved] = 1", connCasper10, 2, 3

              If rs.EOF or weekcount = 0 or weekcount = 6 Then
            %>
            <div style="height: 7em">&nbsp;</div>
            <%
                Else
            %>
            <table style="width: 100%" cellspacing="0" cellpadding="0">
            <%
                    Do While Not rs.EOF
            %>
              <tr>
                <td <%=CheckNumber(count)%>>
                  <a href="javascript:clickView('<%=rs("ID")%>')" style="width: 100%"><%=rs("RequestedBy") & iif(Not rs("Reason") = "", " - " & rs("Reason"), "") & iif(rs("Duration") = "Partial", " - " & rs("DepartureTime"), "")%></a>
                </td>
              </tr>
            <%
                      If count = 2 Then
                        count = 1
                      Else
                        count = count + 1
                      End If
                      rs.MoveNext
                    Loop
                  rs.close
                  count = 1
            %>
            </table>
            <%
                End If
            %>
            </td>
            <%
                range = range + 1
                If weekcount = 6 Then
            %>
          </tr>
          <tr>
            <%
                  weekcount = 0
                Else
                  weekcount = weekcount + 1
                End If
                dayNum = dayNum + 1
              Loop
            %>
          </tr>
        </table>
        <br />
        <br />
        <a href="javascript:clickAdd();" class="button noPrint" style="margin: 0 55em">Request&nbsp;Time</a><br /><br /><br />
            <% 
              If Session("ha_Permissions") = "99" Then
                Set rs = Server.CreateObject("ADODB.Recordset")
                rs.Open "SELECT * FROM ha_EmpTimeRequest WHERE  [IsApproved] IS NULL", connCasper10, 2, 3

                  If Not rs.EOF Then
            %>
        <table style="width: 80%; margin: 0 auto; text-align: center" class="noPrint">
          <tr>
            <th>Requests Received</th>
          </tr>
            <%
               Do While Not rs.EOF
            %>
          <tr>
            <td>
              <a href="javascript:clickEdit('<%=rs("ID")%>');"><%=rs("RequestedBy") & " : " & rs("RequestFrom") & iif(Not rs("RequestTo") = "", " - " & rs("RequestTo") , "")%></a>
            </td>
          </tr>
            <%    
                 rs.MoveNext
               Loop
            %>
        </table>
        <br />
        <br />
        <br />
            <%
             End If
          
             rs.Close
           End If
            %>
      </td>
    </tr>
  </table>
  <!--#include virtual="ha_BackOffice/includes/BodyParts/ha_PageTail.inc"-->
  <!--#include virtual="ha_backoffice/includes/datepick.inc"-->
</body>
</html>
<%
  Function DateFormat(aDate)
    ' declare variables
    Dim result, d, m, y
      
    m = Month(aDate)
    d = Day(aDate)
    y = Year(aDate)
    If m < 10 Then
      m = "0" & m
    End If

    If d < 10 Then
      d = "0" & d
    End If
    
    result = y & "-" & m & "-" & d

    DateFormat = result
  End Function
%>