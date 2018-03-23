<%
    Session("Redirect") = ""
    ColorTab = 1
    PageHeading = "Webpage Hits"
%>

<!--#include virtual="includes/connection.asp"-->
<!--#include virtual="ha_backoffice/includes/Config.asp"-->

<html>
<head>
    <title>Harvest American Webpage Hits</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <script src="includes/checkDate.js" type="text/javascript"></script>
    <script src="datepick/ts_picker.js" type="text/javascript"></script>
    
    <link href="<%=strPathIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet">
    <link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>menu/dropdowntabfiles/halfmoontabs.css">

    <script type="text/javascript" src="<%=strPathIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>

    <style type="text/css">
      table
        {
        empty-cells: show;
        borders-collapse: collapse;
        }
      th {font-size: 12px;}
      td {font-size: 12px;}
    </style>

    <script language="JavaScript" src="http://www.HarvestAmerican.info/includes/UserTracking/UserTrackingPost.js"
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

<script language="javascript" type = "text/javascript">

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

function formLoad()
    {
    UserTracking('formLoad','ha_WebpageHits.asp','','');
    }

    </script>

</head>

<body class="gray_desktop" onload="formLoad();">
  <table align="center" cellspacing="0" cellpadding="0" class="MainBody">
    <tr>
      <td>

        <!--#include virtual="ha_Backoffice/includes/bodyparts/ha_PageHeading.inc"-->

            <table width="85%" border="0" cellspacing="0" cellpadding="0" align="center">
              <tr> 
                <td> 
                  <table width="100%" cellspacing="5" cellpadding="5">

                    <tr>
                      <td>
                        <form method="post" action="ha_WebpageHits.asp" name="frm" style="text-align: center">
                          Start Date
                          <input type="Text" name="startdate" value="<%=d1%>" size="10" />
                          <a href="javascript:show_calendar('document.frm.startdate', document.frm.startdate.value);">
                          <img src="datepick/cal.gif" width="16" height="16" border="0" alt="Calendar" /></a>&nbsp;&nbsp;

                          End Date
                          <input type="Text" name="enddate" value="<%=d2%>" size="10" />
                          <a href="javascript:show_calendar('document.frm.enddate', document.frm.enddate.value);">
                          <img src="datepick/cal.gif" width="16" height="16" border="0" alt="Calendar" /></a>&nbsp;&nbsp;

                          <a href="javascript:clickSubmit();" class="greenButton">Go</a>
                          &nbsp;&nbsp;&nbsp;&nbsp;
                          <a href="javascript:changeDates('<%=DateAdd("d",-1,eDate)%>');" class="greenButton">&lt; Previous</a>
                          <a href="javascript:changeDates('<%=date()%>');" class="greenButton">Today</a>
                          <a href="javascript:changeDates('<%=DateAdd("d",1,eDate)%>');" class="greenButton">Next &gt;</a>
                        </form>
                      </td>
                    </tr>

                  </table>
                </td>
              </tr>

              <tr>
                <td>
                  <table width="100%" cellspacing="5" cellpadding="5" align="center">
                    <tr bgcolor="#008000">
                      <th width="45%" style="color: White">Page</th>
                      <th width="25%" style="color: White">Event</th>
                      <th width="15%" style="color: White">Count</th>
                      <th width="15%" style="color: White">Distinct IP</th>
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
                              strColor = "#B4D7BF"
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
                      </td>
                      <td align="right">
                        <%=rs("DistinctIP")%>
                      </td>
                    </tr>

            <%
                            rs.MoveNext
                        loop
                    end if
                    rs.close
                    set rs = nothing
            %>

                    <tr bgcolor="#008000">
                    <td style="color: White">
                        <b>Different Page Count:</b>&nbsp;
                    <%=recordCount%>
                    </td>
                    <td>
                      &nbsp;
                    </td>
                    <td align="right" style="color: White">
                        <b>Total:</b>&nbsp;
                    <%=Total%>
                    </td>
                    <td align="right" style="color: White">
                        <b>Total IP:</b>&nbsp;
                    <%=TotalIP%>
                    </td>
                  </tr>
                  <tr>
                    <td height="10">
                    </td>
                  </tr>
                  </table>
                </td>
              </tr>

            </table>
      </td>
    </tr>
  </table>
  <!--#include virtual="ha_BackOffice/includes/BodyParts/ha_PageTail.inc"-->
</body>
</html>
