<?xml version="1.0" encoding="ISO-8859-1"?>
<!DOCTYPE html PUBLIC "-//WAPFORUM//DTD XHTML Mobile 1.0//EN" "http://www.wapforum.org/DTD/xhtml-mobile10.dtd">

<!--#include virtual="includes/connection.asp"-->

<html xmlns="http://www.w3.org/1999/xhtml">
  <head>
    <title>Harvest American, Inc. | Today's Sales</title>
    <link href="includes/mobile.css" rel="stylesheet" type="text/css">

    <% 
    Response.AddHeader "Refresh", "60" 
    %>

  </head>
  <body>
  
<%

    Dim rs
    Dim SQL

    if isnull(Request("wDate")) or Request("wDate") = "" then
        wDate = Date()
    else
        wDate = Request("wDate")
    end if


    set rs = server.createObject("ADODB.Recordset")
        SQL = "ni_DailyActivity('" & FormatDateTime(wDate,2) & _
        "', '" & FormatDateTime(wDate,2) & "')"

    rs.open sql,objcon,3,3,4

%>

<!--#include virtual="mobile/includes/header.asp"-->

<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td>

      <table width="100%" border="0" cellspacing="1" cellpadding="1">
        <tr>
          <td align="left">
            <a href="todayssales.asp?wDate=<%=dateadd("d",-1,wDate)%>" class="button">Prev</a>
          </td>
          <td align="center">
            <small><b><%=wDate%></b></small>
          </td>
          <td align="right">
            <a href="todayssales.asp?wDate=<%=dateadd("d",1,wDate)%>" class="button">Next</a>
          </td>
        </tr>
      </table>

    </td>
  </tr>

  <tr><td>&nbsp;</td></tr>

  <tr>
    <td>

      <table class="padsales" width="180" border="0" cellspacing="1" cellpadding="1" align="left">
        <tr>
           <td align="left" width="30">
              <span class="company">Number-it</span>
           </td>
           <td align="right" width="70">
              <%=FormatCurrency(rs("salesN"),2)%><br />
           </td>
        </tr>
        <tr>
           <td align="left">
              <span class="company">Harvest</span>
           </td>
           <td align="right">
              <%=FormatCurrency(rs("salesH"),2)%><br />
           </td>
        </tr>
        <tr>
           <td align="left">
              <span class="company">Zip2Tax</span>
           </td>
           <td align="right">
              <%=FormatCurrency(rs("salesZ"),2)%><br />
           </td>
        </tr>
        <tr>
           <td align="left">
              <span class="company">Total</span>
           </td>
           <td align="right">
              <%=FormatCurrency(rs("salesT"),2)%><br />
           </td>
        </tr>
      </table>

    </td>
  </tr>

  <tr><td>&nbsp;</td></tr>

  <tr>
    <td>

      <table width="100%" border="0" cellspacing="1" cellpadding="1">
        <tr>
          <td align="left">
            <a href="https://www.harvestamerican.org/dave/Index.asp" class="button">Menu</a>
          </td>
          <td align="left">
            <a href="http://www.number-it.com/Home/BackOffice/boDailyActivity.asp" class="button">D.Activity</a>
          </td>
          <td width="100%" align="left">
            &nbsp;
          </td>
        </tr>
      </table>

    </td>
  </tr>
</table>
        
   
<!--#include virtual="mobile/includes/footer.inc"-->

  </body>
</html>