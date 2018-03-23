<html>

<!--#include virtual="includes/database.inc"-->


<head>
    <title>Harvest American Statistics</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <script 'src="includes/checkDate.js" type="text/javascript"></script>
    <script 'src="datepick/ts_picker.js" type="text/javascript"></script>
    <link href="includes/HarvestAmerican.css" type="text/css" rel="stylesheet">
    <script 'src="http://www.HarvestAmerican.info/includes/UserTracking/UserTrackingPost.js"
      type="text/javascript"></script>

<%
     Response.buffer=true
     dim d1
     dim d2
     dim eDate

     if isnull(Request("StartDate")) or Request("StartDate") = "" then
        d1 = date()
     else
        d1 = Request("StartDate")
     end if

     if isnull(Request("EndDate")) or Request("EndDate") = "" then
        d2 = date()
     else
        d2 = Request("EndDate")
     end if
     eDate = d2
%>

<script 'type = "text/javascript">

    var isIE = 1;
    if (navigator.appName == "Netscape")
        { 
        isIE = 0;
        }

    function changeDates(d)
        {
        document.frm.startdate.value = d;
        document.frm.enddate.value = d;
        clickSubmit();
        }

    function checkStartDate()
        {
        var d = document.frm.startdate;
        if (d.value.length != 0) 
           {
           if (isDate(d.value)==false)
              {
                d.focus()
                return false
              }
            return true
            }
            alert('Please enter a start date');
            d.focus()
            return false
          }    

        function checkEndDate()
          {
          var d = document.frm.enddate;
          if (d.value.length != 0) 
            {
            if (isDate(d.value)==false)
              {
                d.focus()
                return false
              }
            return true
            }
            alert('Please enter an end date');
            d.focus()
            return false
          }    

        function clickSubmit()
          {
          if (checkStartDate() && checkEndDate())
            {
            document.frm.submit();
            }
          }

function altOn(obj, t, b, e)
    {
    var altScreen = document.getElementById('mouseoverDiv');
    var sTop = document.body.scrollTop;
    var cursor = getPosition(e);

    altScreen.style.top = 100 + sTop;
    altScreen.style.top = cursor.y - 20;
    altScreen.style.left = cursor.x - 600;

    document.getElementById('mouseoverTitle').innerHTML = t;
    b = b.replace(/~/g,"<br>");
    document.getElementById('mouseoverBody').innerHTML = b;

    altScreen.style.visibility = 'visible';
    }

function altOff()
    {
    document.getElementById('mouseoverDiv').style.visibility = 'hidden';
    }

function getPosition(e) {
    //Firefox event gets passed in, but IE needs to get it now
    if (!e) var e = window.event;
    var cursor = {x:0, y:0};
    if (e.pageX || e.pageY) {
        cursor.x = e.pageX;
        cursor.y = e.pageY;
    } 
    else {
        cursor.x = e.clientX + 
            (document.documentElement.scrollLeft || 
            document.body.scrollLeft) - 
            document.documentElement.clientLeft;
        cursor.y = e.clientY + 
            (document.documentElement.scrollTop || 
            document.body.scrollTop) - 
            document.documentElement.clientTop;
    }
    return cursor;
}


function formLoad()
    {
    UserTracking('formLoad','ha_Stats.asp','','');
    }

    </script>

</head>

<body bgcolor="#FFFFFF" onLoad="formLoad();">
<table width="100%" border="0" cellspacing="0" cellpadding="0">

  <tr> 
    <td align="center"> 
      <font size="4">
        <B>Harvest American Statistics</B><br><br>
      </font>
    </td>
  </tr>

  <tr> 
    <td> 
      <table width="100%" border="1" cellspacing="0" cellpadding="2" BorderColor="White">

        <tr>
          <td>
      <form method="Post" Action="ha_Stats.asp" name="frm">
        Start Date
        <input type="Text" name="startdate" value="<%=d1%>" size="10">
        <a href="javascript:show_calendar('document.frm.startdate', document.frm.startdate.value);">
        <img src="datepick/cal.gif" width="16" height="16" border="0" alt="Calendar"></a>&nbsp;&nbsp;

        End Date
        <input type="Text" name="enddate" value="<%=d2%>" size="10">
        <a href="javascript:show_calendar('document.frm.enddate', document.frm.enddate.value);">
        <img src="datepick/cal.gif" width="16" height="16" border="0" alt="Calendar"></a>&nbsp;&nbsp;

        <a href="javascript:clickSubmit();" class="button">Go</a>
        &nbsp;&nbsp;&nbsp;&nbsp;
        <a href="javascript:changeDates('<%=DateAdd("d",-1,eDate)%>');" class="button2">< Previous</a>
        <a href="javascript:changeDates('<%=date()%>');" class="button2">Today</a>
        <a href="javascript:changeDates('<%=DateAdd("d",1,eDate)%>');" class="button2">Next ></a>
      </form>
          </td>
        </tr>

      </table>
    </td>
  </tr>

  <tr>
    <td>
      <table width="100%" border="1" cellspacing="0" cellpadding="2" BorderColor="White">
	    <tr bgcolor="#80FF80">
	  	  <th width="30%">Page</th>
		  <th width="30%">Event</th>
		  <th width="20%">Count</th>
		  <th width="20%">Distinct IP</th>
	    </tr>

<%
	    set rs=server.createObject("ADODB.Recordset")
          rs.open "ha_TrackingStats('" & d1 & "','" & d2 & "')",objcon,3,3,4

	    if not RS.EOF then
		    do while not rs.eof
		        recordCount = recordCount + 1
		        LineCount = LineCount + 1
		        If LineCount > 2 then
			      LineCount = 0
			      strColor = "#C8FFC8"
		        Else
			      strColor = "#FFFFFF"
		        End If
                    Total = Total + cint(rs("Count"))
                    TotalIP = TotalIP + cint(rs("DistinctIP"))
%>

		<tr bgcolor=<%=strColor%>>
		  <td align="left">
                    <%=rs("Page")%>
		  </td>
		  <td>
		    <%=rs("Operation")%>
		  </td>
		  <td align="right">
		    <%=rs("Count")%>
                                    <span id = 'count1'>
                                      <img src='includes/ViewCount.gif'
                                        border='0' alt='View Count'
                                        onMouseOver='altOn(this, "Count", "BodyText", event);'
                                        onMouseOut='altOff();' />
                                    </span>
		  </td>
		  <td align="right">
		    <%=rs("DistinctIP")%>
                                    <span id = 'count2'>
                                      <img src='includes/ViewCount.gif'
                                        border='0' alt='View Count'
                                        onMouseOver='altOn(this, "Distinct IP", "BodyText", event);'
                                        onMouseOut='altOff();' />
                                    </span>
		  </td>
	    </tr>

<%
	            rs.MoveNext
	        loop
        end if
        rs.close
        set rs = nothing
%>

        <tr bgcolor="#80FF80">
	    <td>
            <b>Different Page Count:</b>&nbsp;
		<%=recordCount%>
	    </td>
          <td>
            &nbsp;
          </td>
	    <td align="right">
            <b>Total:</b>&nbsp;
		<%=Total%>
	    </td>
	    <td align="right">
            <b>Total IP:</b>&nbsp;
		<%=TotalIP%>
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
<!--#include virtual='ControlPanel/stats/includes/mouseoverDiv.inc'-->
</body>
</html>
