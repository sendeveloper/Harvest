<!DOCTYPE html>

  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsPermissions.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsMouseover.inc"-->  
  
<%
	Title = "Harvest American Calendar"
	
	Dim rs: Set rs = Server.CreateObject("ADODB.Recordset")
	Call mouseoverDivAbove(): 'Creates the mouseover div

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

<html>
<head>
  <title><%=Title%></title>

  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
  
  <link href="/ha_Backoffice/includes/datepick/jquery-ui.css" rel="stylesheet" />
  <link href="/ha_Backoffice/includes/ha_BackOfficePopup.css" type="text/css" rel="stylesheet">
  <script language="javascript" type="text/javascript" src="/includes/haFunctions.js"></script>
  
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
		
	  h1
		{
		font-size:		16pt;
		}
  </style>
	
  <style media="print" type="text/css">		
	.noPrint
		{
		display: none !important;
		}
  </style>

  <script type="text/javascript">
  
    function clickChange(date) 
		{
        window.location = 'ha_Calendar.asp?date=' + date;
        }		
		
	function formLoad()
		{
		SetScreen(1000,900);
		}
		
  </script>
</head>

<body onLoad="formLoad();">

	<span class="popupHeading"><%=Title%></span>
	  
  <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">	  
	<tr>
      <td>
	
        <table style="width: 920px; margin: 0 auto;" cellpadding="0" cellspacing="2">
          <tr>
            <td width="35%">
              <a class="buttonGray button50 noPrint" 
			    href="ha_Calendar.asp?date=<%=MakeDate(yearCur - 1, monthCur)%>">< <%=yearCur - 1%></a>&nbsp;
			  <a class="buttonGray button50 noPrint" 
			    href="ha_Calendar.asp?date=<%=MakeDate(yearPrev, monthPrev)%>">< <%=MonthName(monthPrev, True)%></a>
			</td>
			<td align="center" width="20%">
			  <h1><%=MonthName(monthCur) & " " & yearCur %></h1>
			</td>
			<td align="right" width="35%">
			  <input type="date" name="changedate" id="changedate" class="datepick noPrint" value="<%=datePH%>" />
				<a class="buttonGray noPrint" href="javascript:void(0)" onclick="clickChange(changedate.value)">Go</a>&emsp;&emsp;
				<a class="buttonGray button50 noPrint" 
				  href="ha_Calendar.asp?date=<%=MakeDate(yearNext, monthNext)%>"><%=MonthName(monthNext, True)%> ></a>&nbsp;
				<a class="buttonGray button50 noPrint" 
				  href="ha_Calendar.asp?date=<%=MakeDate(yearCur + 1, monthCur)%>"><%=yearCur + 1%> ></a>
			</td>
          </tr>
		</table>

        <table style="width: 920px; margin: 0 auto;" cellpadding="0" cellspacing="2">

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
			SQL = "ha_Employee_Calendar_read('" & datePH & "')"
			rs.Open SQL, connLocal, 3, 3, 4
			
            If rs.EOF or weekcount = 0 or weekcount = 6 Then
%>
            <div style="height: 5em">&nbsp;</div>
<%
			Else
%>
            <table style="width: 100%" cellspacing="0" cellpadding="0">
			
<%
				Do While Not rs.EOF				
					Select Case rs("Type")
					Case "Event"
						TitleText = "Event"
						BodyText = rs("Reason")
						DisplayStyle = "color: red; font-weight: bold; text-align: center;"
					Case "Holiday"
						TitleText = "Holiday"
						BodyText = rs("Reason")
						DisplayStyle = "color: red; font-weight: bold; text-align: center;"
					Case Else
						If rs("Duration") = "All Day" Then
							TitleText = rs("RequestFor") & " Out"
						Else
							TitleText = rs("RequestFor") & " Partial"
						End If
						BodyText = ""
						If rs("Reason") > "" Then
							BodyText = BodyText & Replace(rs("Reason"),"'","\'") & "<br>"
						End If					  
						If rs("DepartureTime") > "" Then
							BodyText = BodyText & "Leaving at " & rs("DepartureTime")
						End If
						DisplayStyle = ""
					End Select
%>

              <tr valign="top">
                <td style="<%=DisplayStyle%>"
				  &emsp;
				  onMouseOver="altOnAbove(this, '<%=TitleText%>', '<%=rs("DisplayDate")%>', '<%=BodyText%>', event);"
				  onMouseOut="altOffAbove();">				
				  <%=TitleText%>
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
		
      </td>
	</tr>
  </table>

    <span class="popupButtons noPrint" style="padding-bottom: 3em;">
      <a href="javascript:window.close();" class="bo_Button80">Close</a>
    </span>

</body>
</html>
