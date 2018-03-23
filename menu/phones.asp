<html>

<!--#include virtual="includes/connection.asp"-->
<!--#include virtual="menu/includes/phone-functions.asp"-->

<head>
    <title>Harvest American Phone Call Log</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <script language="JavaScript" src="includes/checkDate.js" type="text/javascript"></script>
    <script language="JavaScript" src="datepick/ts_picker.js" type="text/javascript"></script>
    <script language="JavaScript" src="<%=strPathIncludes%>haFunctions.js"></script>

    <link href="<%=strPathControlPanel%>includes/HarvestAmerican.css" type="text/css" rel="stylesheet">

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
</head>

<body>
<table width="100%" align="center" border="0" cellspacing="0" cellpadding="0">

  <tr> 
    <td width="80%" align="left"> 
      <h1>Harvest American Phone Call Log</h1>
    </td>
    <td width="20%" align="right">
      <a href="/menu/index.asp" class="button">Back</a>
    </td>
  </tr>

  <tr>
    <td colspan="2">
<%
          recordCount = 0
          LineCount = 0

          set rs=server.createObject("ADODB.Recordset")
          rs.open "ha_Phone_CallLog_view", objcon, 3, 3, 4

'------------------------------------Starting Paging from Here-----------------------------------
	iRecordperpage=15
	rs.PageSize = iRecordperpage
	nRecordsPerPage = rs.PageSize
	rs.CacheSize = rs.PageSize
	intPageCount = rs.PageCount 
	intRecordCount = rs.RecordCount

        icurrentPage = TRIM(REQUEST("page"))
        if icurrentPage = "" then 
            icurrentPage = 0
        else
            icurrentPage = CInt(request("page"))
        end if

        If CInt(icurrentPage) > CInt(intPageCount) Then icurrentPage = intPageCount
        If CInt(icurrentPage) <= 0 Then intPage = 1
			
%>
<%      If intRecordCount > 0 Then %>
      <table width="100%" align="left" border="1" cellspacing="0" cellpadding="2">
        <tr bgcolor="#FFFF66">
          <th style="width: 1em">#</th>
          <th style="width: 16em">Call Date</th>
          <th style="width: 5em">Extension</th>
          <th style="width: 4em">CO</th>
          <th style="width: 1em">Incoming</th>
          <th style="width: 50em">Dial Number</th>
          <th style="width: 4em">Ring</th>
          <th style="width: 8em">Duration</th>
          <!-- <th style="width: 11em">Cost</th> -->
          <!-- <th style="width: 10em">Access Code</th> -->
          <!-- <th style="width: 3em">CD</th> -->
          
          <!-- <th style="width: 50%">&nbsp;</th> -->
        </tr>
<%
            rs.AbsolutePage = icurrentPage + 1
            intStart = rs.AbsolutePosition 
            If CInt(icurrentPage) = CInt(intPageCount)-1 Then
                intFinish = intRecordCount
            Else
                intFinish = intStart + (rs.PageSize - 1)
            End if

            n = 1
            Do While Not rs.eof
              Do While n <= iRecordperpage and Not rs.eof
                'on error resume next
                     recordCount = recordCount + 1
                  LineCount = LineCount + 1
                  If LineCount > 2 then
                    LineCount = 0
                    strColor = "#FFFFCC"
                  Else
                    strColor = "#FFFFFF"
                  End If

%>

        <tr bgcolor=<%=strColor%>>
          <td style="border-color: #EEEEEE;">
            <%=rs("num")%>
          </td>
          <td style="border-color: #EEEEEE;">
            <%=rs("CallDate")%>
          </td>
          <td style="border-color: #EEEEEE;">
            <%=rs("Ext")%>
          </td>
          <td style="border-color: #EEEEEE;">
            <%=rs("CO")%>
          </td>
          <td style="border-color: #EEEEEE;">
            <% If rs("Incoming") = "True" Then %>
              I
            <% Else %>
              &nbsp;
            <% End If %>
          </td>
          <td style="border-color: #EEEEEE;">
            <%=rs("DialNumber")%>
          </td>
          <td style="border-color: #EEEEEE;">
            <% If Trim(rs("Ring")) = "" Then %>
              &nbsp;
            <% Else %> 
              <%=rs("Ring")%>
            <% End If %>
          </td>
          <td style="border-color: #EEEEEE;">
            <%=rs("Duration")%>
          </td>
          <!--
          <td style="border-color: #EEEEEE;">
            <%=rs("Cost")%>
          </td>
          -->
          <!--
          <td style="border-color: #EEEEEE;">
            <%=rs("AccCode")%>
          </td>
          -->
          <!--
          <td style="border-color: #EEEEEE;">
            <%=rs("CD")%>
          </td>
          -->
          <!-- 
          <td style="border-color: #EEEEEE;">
            &nbsp;
          </td> 
          -->
        </tr>
<%
                    rs.MoveNext
                    n=n+1
                    If rs.eof Then 
                      Exit Do
                    End If
                 Loop
                 i=i+1
                 rs.movenext
            Loop


            If i < iRecordperpage Then
              response.write "<tr><td height='100%'>&nbsp;</td></tr>"
            End If  
'        End If
        rs.close
        set rs = nothing
        objcon.close
%>

        <tr bgcolor="#FFFF66">
          <td>
            <b>Calls:</b>&nbsp;
            <%=intStart%>&nbsp;-&nbsp;<%=intFinish%>&nbsp;/&nbsp;<%=recordCount%>
          </td>
          <td>
            &nbsp;
          </td>
          <td>
            &nbsp;
          </td>
          <td>
            &nbsp;
          </td>
          <td>
            &nbsp;
          </td>
          <td>
            &nbsp;
          </td>
          <td>
            &nbsp;
          </td>
          <td>
            &nbsp;
          </td>
        </tr>
      </table>
<%
        Else
    'on error resume next
%>
          <td align="center" style="color: C0C0C0; font-weight: bold;">No records found matching that search criteria</td>
<%      End If %>
<div style="border: red solid 5px; ">
  Here is the box:<br><br><br><br><br>
  <%=GetHitCountAndPageLinks(Request.ServerVariables("URL"), intRecordCount, nRecordsPerPage, intPageCount, icurrentPage,"" )%>
</div>
    </td>
    <td>

<!-- end paging -->
    </td>
  </tr>
  
</table>
</body>
</html>
