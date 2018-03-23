<!DOCTYPE html>
<html>
<head>
  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/Config.asp"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsMouseover.inc"-->
  
<%
	Session("Redirect") = ""
	ColorTab = 4
	PageHeading = "Vacation Calendar"

    'Set up date variables
    Dim yearCur, datePH, monthCur, monthPrev, monthNext, assembleDate, dayNum: dayNum = 1

    If Request("date") = "" Then
      monthCur = Month(Now)
      yearCur = Year(Now)
	  datePH = MakeDate(yearCur, monthCur)
    Else
      monthCur = Month(Request("date"))
      yearCur = Year(Request("date"))
	  datePH = MakeDate(yearCur, monthCur)
    End If

    If monthCur = 1 Then
      monthPrev = 12
	  yearPrev = yearCur - 1
    Else
      monthPrev = monthCur - 1
	  yearPrev = yearCur
    End If
	
    If monthCur = 12 Then
      monthNext = 1
	  yearNext = yearCur + 1
    Else
      monthNext = monthCur + 1
	  yearNext = yearCur
    End If
	
	Function MakeDate(y, m)
		MakeDate = cstr(y) & "-" & right("00" & cstr(m),2) & "-01" 
	End Function	
	
	'Determine Number of Days
	If monthCur = 4 Or monthCur = 6 Or monthCur = 9 Or monthCur = 11 Then
	  NOD = 30
	ElseIf monthCur = 2 Then
	  If monthCur Mod 4 = 0 Then
		NOD = 29
	  Else
		NOD = 28
	  End If
	Else
	  NOD = 31
	End If

%>

    <title>Harvest American - <%=PageHeading%></title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <link href="<%=strPathIncludes %>datepick/jquery-ui.css" rel="stylesheet" />
    <link href="<%=strPathIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet" />
    <link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>menu/dropdowntabfiles/halfmoontabs.css" />
    <script type="text/javascript" src="<%=strPathIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>
    <script type="text/javascript" src="<%=strPathBOIncludes%>haFunctions.js"></script>
		
<script type="text/javascript">
    function clickAddVacationRequest()
        {
        var URL = 'ha_viewrequest.asp' +
            '?rMode=Add';
        openPopUp(URL);
        }
	
        function clickView(id) {
            var URL = 'ha_viewrequest.asp?rMode=View&id=' + id;
            window.open(URL, '_blank', 'width=425; height=500');
        }
        function clickEdit(id) {
            var URL = 'ha_viewrequest.asp?rMode=Edit&id=' + id;
            window.open(URL, '_blank', 'width=425; height=500');
        }
        function clickChange(date) {
            window.location = 'ha_Calendar.asp?date=' + date;
        }		
</script>

    <style type="text/css">
        tr.calrowhead
			{
            background-color: #308F5A;
			}
			
        th.calhead
			{
            width: 14.3%;
            color: white;
            border: 5px solid #308F5A;
			}
        
        td.calcell
			{
            border: 2px solid #D8F0E3;
            vertical-align: top;
			padding: 5px;
			}
		
		.calDate
			{
			font-family:	"Times New Roman", Georgia, Serif;
			font-weight: 	bold;
			font-size:		16pt;
            text-align:		right;
			display:		block;
			padding:		2px;
			}		
    </style>
	
    <style media="print" type="text/css">		
		.noPrint
			{
			display: none !important;
			}
	</style>
</head>

<body class="gray_desktop">
  <table cellpadding="0" cellspacing="0" class="MainBody">
    <tr class="noPrint">
      <td>
        <!--#include virtual="ha_Backoffice/includes/bodyparts/ha_PageHeading.inc"-->
      </td>
    </tr>
    <tr>
      <td>
        <table style="width: 1150px; margin: 0 auto;" cellpadding="0" cellspacing="2">
          <tr>
            <td width="35%">
              <a class="buttonLightGreen button50 noPrint" 
			    href="ha_Calendar.asp?date=<%=MakeDate(yearCur - 1, monthCur)%>"><%=yearCur - 1%></a>&nbsp;
			  <a class="buttonLightGreen button50 noPrint" 
			    href="ha_Calendar.asp?date=<%=MakeDate(yearPrev, monthPrev)%>">< <%=MonthName(monthPrev, True)%></a>
			</td>
			<td align="center" width="20%">
			  <h1><%=MonthName(monthCur) & " " & yearCur %></h1>
			</td>
			<td align="right" width="35%">
			  <input type="date" name="changedate" id="changedate" class="datepick noPrint" value="<%=datePH%>" />
				<a class="buttonLightGreen noPrint" href="javascript:void(0)" onclick="clickChange(changedate.value)">Go</a>&emsp;&emsp;&emsp;&emsp;
				<a class="buttonLightGreen button50 noPrint" 
				  href="ha_Calendar.asp?date=<%=MakeDate(yearNext, monthNext)%>"><%=MonthName(monthNext, True)%> ></a>&nbsp;
				<a class="buttonLightGreen button50 noPrint" 
				  href="ha_Calendar.asp?date=<%=MakeDate(yearCur + 1, monthCur)%>"><%=yearCur + 1%> ></a>
			</td>
          </tr>
		</table>
		  
        <table style="width: 1150px; margin: 0 auto;" cellpadding="0" cellspacing="2">
		  <span class="noPrint">
		    <%=TopRightLinks(7, "", "", "", "Add Vacation Request")%>
		  </span>

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
	Dim doW, weekcount, range
		range = 0
		weekcount = 0

        doW = Weekday(datePH) - 1

        Do While Not range = doW
            Response.write "<td class='calcell'>&nbsp;</td>"
			range = range + 1
		Loop

		weekcount = range
		range = 1
              
		Do While Not range = NOD +1
            Response.write "<td class='calcell'><span class='calDate'>" & range & "</span><br>"
            datePH = CDate(monthCur & "-" & dayNum & "-" & yearCur)
            Set rs = Server.CreateObject("ADODB.Recordset")
			SQL = "ha_Employee_Calendar_read('" & datePH & "')"
			rs.Open SQL, connLocal, 3, 3, 4
			
            If rs.EOF or weekcount = 0 or weekcount = 6 Then
%>
            <div style="height: 7em">&nbsp;</div>
<%
			Else
%>
            <table style="width: 100%" cellspacing="0" cellpadding="0">
			
<%
				Do While Not rs.EOF
					TitleText = "Request by " & rs("RequestedBy")
					BodyText = ""
					If rs("Reason") > "" Then
						BodyText = BodyText & Replace(rs("Reason"),"'","\'") & "<br>"
					End If					  
					If rs("Duration") = "Partial" Then
						BodyText = BodyText & "Leaving at " & rs("DepartureTime")
					End If
%>

              <tr valign="top">
                <td
				  &emsp;
				  onMouseOver="altOn(this, '<%=TitleText%>', '<%=datePH%>', '<%=BodyText%>', event);"
				  onMouseOut="altOff();">
				
                  <a href="javascript:clickView('<%=rs("ID")%>')" style="width: 100%">
				    <%=rs("RequestedBy")%></a>
                </td>
              </tr>

<%
					rs.MoveNext
				Loop
%>
            </table>
<%
			End If
			rs.close
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
