<?xml version="1.0" encoding="ISO-8859-1"?>
<!DOCTYPE html PUBLIC "-//WAPFORUM//DTD XHTML Mobile 1.0//EN" "http://www.wapforum.org/DTD/xhtml-mobile10.dtd">

<%
'Response.write Request.ServerVariables("HTTP_USER_AGENT")
'Response.end
%>

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

    If isnull(Request("wDate")) Or Request("wDate") = "" Then
        wDate = Date()
    Else
        wDate = Request("wDate")
    End If

    Dim connBarley: Set connBarley = Server.CreateObject("ADODB.Connection")
	'connBarley.open("driver=SQL Server;server=barley1.harvestamerican.net;uid=davewj2o;pwd=get2it;database=ha_prod")
	connBarley.open("driver=SQL Server;server=66.119.50.228,7843;uid=davewj2o;pwd=get2it;database=ha_prod")
    Set rs = server.createObject("ADODB.Recordset")
        SQL = "ni_DailyActivity('" & FormatDateTime(wDate,2) & _
			"', '" & FormatDateTime(wDate,2) & "', 'mobile')"

    rs.open SQL, connBarley, 3, 3, 4

    If Not rs.eof Then
        SalesN = rs("SalesN")
        SalesW = rs("SalesW")
        SalesH = rs("SalesH")
        SalesZ = rs("SalesZ")
        SalesT = rs("SalesT")
    Else
        SalesN = 0
        SalesW = 0
        SalesH = 0
        SalesZ = 0
        SalesT = 0
    End If

%>

<a href="todayssales.asp?wDate=<%=date%>" title="Today">
<!--#include virtual="mobile/includes/header.asp"-->
</a>
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

      <table class="padsales" width="180" border="0" cellspacing="0" cellpadding="1" align="left">
        <tr>
           <td align="left" width="30">
              <span class="company">Number-it</span>
           </td>
           <td align="right" width="70">
              <%=FormatCurrency(salesN,2)%><br />
           </td>
        </tr>
        <tr>
           <td align="left">
              <span class="company">Wholesale</span>
           </td>
           <td align="right">
              <%=FormatCurrency(salesW,2)%><br />
           </td>
        </tr>
        <tr>
           <td align="left">
              <span class="company">Harvest</span>
           </td>
           <td align="right">
              <%=FormatCurrency(salesH,2)%><br />
           </td>
        </tr>
        <tr>
           <td align="left">
              <span class="company">Zip2Tax</span>
           </td>
           <td align="right">
              <%=FormatCurrency(salesZ,2)%><br />
           </td>
        </tr>
        <tr>
           <td align="left">
              <span class="company">Total</span>
           </td>
           <td align="right">
              <%=FormatCurrency(salesT,2)%><br />
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
            <a href="/menu/dave/Index.asp" class="button">Menu</a>
          </td>
          <td align="left">
            <a href="http://www.number-it.com/Home/BackOffice/boSales/boDailyActivity.asp?lname=davewj2o&pass=)^v(" class="button">Sales</a>
          </td>
          <td width="80%" align="left">
            &nbsp;
          </td>
        </tr>
      </table>

    </td>
  </tr>

  <tr><td>&nbsp;</td></tr>

</table>
        
   
<!--#include virtual="mobile/includes/footer.inc"-->

  </body>
</html>
