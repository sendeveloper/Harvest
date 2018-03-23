<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<%
     ThisPage = "https://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL")
%>

<!--#include virtual="includes/connection.asp"-->

<%
  If session("LoggedIn")<>"True" Or isNull(session("LoggedIn")) Then
    Response.Redirect strPathIncludes & "Login.asp"
  End If
%>

<html xmlns="http://www.w3.org/1999/xhtml">
  <head>
    <title>Harvest American, Inc. | <%=trim(Session("ha_fname"))%>'s Time Card</title>
    
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <link href="<%=strPathMenuIncludes%>ha_menu.css" type="text/css" rel="stylesheet">
    
    <script type="text/javascript" src="<%=strPathMenuIncludes%>js/functions.js"></script>
  </head>

  <body>
    <!--#include virtual="menu/includes/ha_header.inc"-->
    <!--#include virtual="menu/includes/ha_menu.inc"-->

    <div id="content">
      <div id="columnA">
<%
  Dim rs
  'Dim sqlText
  Dim rs2
  Dim SQL2
                        
  NewYears = DateSerial(Year(Date), 1, 1)

  If Request.QueryString("EmpID") <> 0 Then
    Set rs = server.createObject("ADODB.Recordset")
    sqlText="SELECT top 27 * FROM ha_EmpHours WHERE EmpID='" & Request.QueryString("EmpID") & "' ORDER BY DateWorked DESC"
    rs.open sqlText,objcon,2,3
  Else                
    If Session("ha_EmpID") = "" Then
        Set rs = server.createObject("ADODB.Recordset")
        sqlText="SELECT * FROM ha_EmpHours ORDER BY EmpID,DateWorked DESC"
    
        rs.open sqlText,objcon,2,3

      Else
        Set rs = server.createObject("ADODB.Recordset")
        sqlText="SELECT top 27 * FROM ha_EmpHours WHERE EmpID='"&Session("ha_EmpID")&"' ORDER BY DateWorked DESC"
        rs.open sqlText,objcon,2,3
      End If
  End If
%>

    <table width="440" border="0" cellspacing="1" cellpadding="0" align="center">
      <tr>
        <td colspan="2">
          <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
            <tr>
              <td>
                
                <table style='color:#336600;' width="100%" height="50" border="0" cellspacing="1" cellpadding="3" align="center">
                  
                  <tr>
                    <th width='20%'>
                      Date
                    </th>
                    <th width='12%'>
                      Hours
                    </th>
                    <th width='18%'>
                      Type
                    </th>
                    <th width='20%'>
                      Date Paid
                    </th>
                    <th width='15%'>
                      &nbsp;
                    </th>
                    <th width='15%'>
                      &nbsp;
                    </th>
                  </tr>
                  <tr>
                    <td colspan='6'>
                      <hr style='color:#CDE2A1;'>
                    </td>
                  </tr>
                  
                  
                  
                  <table width="100%" border="0" cellspacing="0" cellpadding="3" align="center">
                    
<%
  Dim iRecordCount
  iRecordCount = 0
   
  If not rs.eof Then
    Do While Not rs.eof
      If iRecordCount Mod 2 = 0 Then
%>
                    <tr class="rowcolor" height="29">
<%
      Else
%>
                    <tr height="29">
<%
      End If
%>

                      <td width='20%' align='right'>
                        <%=rs("DateWorked")%>
                      </td>

                      <td width='12%' align='center'>
                        <%=rs("QuantityHours")%>
                      </td>
 
                      <td width='18%' align='center'>
<%
      Select Case rs("HoursType")
      Case 1
        response.write("Regular")
      Case 2
        response.write("Vacation")
      Case 3
         response.write("Sick Time")
      Case 4
         response.write("Holiday")
      End Select
%>
                      <td width='20%' align='right'>
                        <%=rs("DatePaid")%>
                      </td>
<%
      If rs("WorkDescription") = "" Or isNull(rs("WorkDescription")) = True Then
%>
                      <td width='15%' align='center'>
                        &nbsp;
                      </td>
<%
      Else
%>
                      <td width='15%' align='right' style='padding-bottom: 4px;'>
                        <a href="javascript: void(0)" onclick="popup('ha_timeclockNotes.asp?id=<%=rs("EmpID")%>')" class="button">Notes</a>
                      </td>
<%
      End If
%>
                      <td width='15%' align='center' style='background: #F5F9EC; padding-bottom: 4px;'>
                        <a href="javascript: void(0)" onclick="popup('ha_timeclockEdit.asp?id=<%=rs("id")%>')" class="button">Edit</a>
                      </td>
                    </tr>
                    
<%
      iRecordCount = iRecordCount + 1
      rs.MoveNext
    Loop
  End If
  rs.Close
  Set rs = Nothing
%>

                  </table>
              </td>
            </tr>
            </table>
        </td>
      </tr>
      </table></div>

      <div id="columnB">
<%
   Set rs2 = server.createObject("ADODB.Recordset")
   SQL2="ha_TimeClockStatistics('" & Session("ha_EmpID") & "')"
   
   rs2.open SQL2,objcon,2,3
%>
        <table width="180" border="0" cellspacing="1" cellpadding="1" align="center">
          <tr>
            <td align="center">
<% If len(Session("ha_fname")) + len(Session("ha_lname")) > 15 Then %>
              <h2 style="font-size: medium;">
<% Else %>
              <h2>
<% End If %>
              <br><%=Session("ha_fname")%>&nbsp;<%=Session("ha_lname")%><br>Time Card</h2><br><br>
            </td>
          </tr>
          <tr>
            <td width='100%' align='center'>
              <a href="javascript: void(0)" onclick="popup('ha_timeclockEdit.asp?cMode=Add')" class="button">ADD RECORD</a>
            </td>
          </tr>
        </table>        
        <table class="outlined" width="180" border="0" cellspacing="0" cellpadding="0" align="center">
          <tr>
            <td align="center" bgcolor="#CDE2A1"><h3>Statistics</h3></td>
          </tr>
<%
  If rs2("VacationTotal") <> "" Or isNull(rs2("VacationTotal")) = False Then
    Session("Vacation") = True

%>
          <tr>
            <td align="center"><h5>Vacation Hours Used<br>
<%
    If rs2("VacationUsed") <> "" Or isNull(rs2("VacationUsed")) = False Then
      Response.Write(rs2("VacationUsed"))
    Else
      Response.Write("0")
    End If
%>
              /<%=rs2("VacationTotal")%></h5>
            </td>
          </tr>
<%
  Else
    Session("Vacation") = False
  End If
  If rs2("PaidTimeOffTotal") <> "" Or isNull(rs2("PaidTimeOffTotal")) = False Then
    Session("PaidTimeOff") = True

%>
          <tr>
            <td align="center"><h5>Sick Time Used<br>
<%
    If rs2("PaidTimeOffUsed") <> "" Or isNull(rs2("PaidTimeOffUsed")) = False Then
      Response.Write(rs2("PaidTimeOffUsed"))
    Else
      Response.Write("0")
    End If
%>
              /<%=rs2("PaidTimeOffTotal")%></h5>
            </td>
          </tr>
<%
    Else
      Session("PaidTimeOff") = False
    End If

    If rs2("HoursTotal") <> "" or IsNull(rs2("HoursTotal")) = False then
%>
          <tr>
            <td align="center"><h5>Total Hours Worked for Harvest American, Inc.<br><%=rs2("HoursTotal")%></h5>
            </td>
          </tr>
<%
  Else
%>
          <tr>
            <td align="center"><h4>No Data Found To Generate Statistics</h4>
            </td>
          </tr>
<%
  End If
%>
        </table>
<%
  rs2.Close
  Set rs2 = Nothing
%>

     <!--#include virtual="menu/includes/ha_domains.inc"-->

      </div>
      <div style="clear: both;">&nbsp;</div>
    </div>
    
    <!--#include virtual="menu/includes/ha_footer.inc"-->
    
  </body>
</html>
