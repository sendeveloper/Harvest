<!DOCTYPE html>
<html>
<head>
  <%
    Session("Redirect") = strPathIncludes & "login.asp"
	ColorTab = 5
	PageHeading = "Timeclock Administration"

    Function CheckNumber(num)
      If num MOD 3 = 0 Then
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
  <link href="http://code.jquery.com/ui/1.10.2/themes/south-street/jquery-ui.css" rel="stylesheet" />
  <link href="<%=strPathIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet" />
  <link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>menu/dropdowntabfiles/halfmoontabs.css" />
  <script type="text/javascript" src="<%=strPathIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>
  <style type="text/css">
        tr.calrowhead
        {
            background-color: #008000;
        }
    </style>
</head>
<body class="gray_desktop">
  <table cellpadding="0" cellspacing="0" class="MainBody" style="margin: 1em auto 0; width: 1200px">
    <tr>
      <td>
        <!--#include virtual="ha_Backoffice/includes/bodyparts/ha_PageHeading.inc"-->
      </td>
    </tr>
    <tr>
      <td>
        <% 
          Dim rs, strsql, count: count = 1
          strsql = "SELECT * FROM ha_EmpAccounts WHERE [permitTimeClock] = 1 AND ([Contractor] = 0 OR [Contractor] IS NULL) ORDER BY [LastName]"
          Set rs = Server.CreateObject("ADODB.Recordset")
          rs.Open strsql, connCasper10, 1, 3
        %>
        <table style="width: 80%; margin: 0 auto" cellpadding="5">
          <tr class="calrowhead">
            <th style="color: white; width: 50%;">Name</th>
            <th style="color: white; width: 50%;">Date Hired</th>
          </tr>
          <% Do While Not rs.EOF %>
          <tr<%=CheckNumber(count) %>>
            <td><a href="ha_timeclock.asp?id=<%=rs("ID")%>"><%=rs("LastName") & ", " & rs("FirstName") %></a></td>
            <td><%=rs("HiredDate")%></td>
          </tr>
          <%   count = count + 1
               rs.MoveNext
             Loop %>
        </table>
      </td>
    </tr>
    <tr>
      <td><br /><br />&nbsp;</td>
    </tr>
  </table>
  <!--#include virtual="ha_BackOffice/includes/BodyParts/ha_PageTail.inc"-->
</body>
</html>
