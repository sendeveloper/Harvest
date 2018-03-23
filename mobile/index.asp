<?xml version="1.0" encoding="ISO-8859-1"?>
<!DOCTYPE html PUBLIC "-//WAPFORUM//DTD XHTML Mobile 1.0//EN" "http://www.wapforum.org/DTD/xhtml-mobile10.dtd">
<!--#include file="./includes/ha_secure.inc"-->
<!--#include file="./includes/connection.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
  <head>
    <title>Today's Sales</title>

    <style type="text/css"> 
        .header  { background: #CDE2A1; color: #336600; border-bottom: solid 2px #336600; margin: 0 0 5px 0; padding: 2px; text-align: center}
        .footer  { background: #CDE2A1; color: #336600; border-top: solid 2px #336600; margin: 10px 0 0 0; text-align: center}
        .company { font-weight: bold; color: #336600}
        .content { text-align: center}
        hr       { clear: both; border:solid; border-width:1px; border-bottom-color:#007300; border-top-color:#ffffff; border-left-color:#ffffff; border-right-color:#ffffff;}
        .top-news img { float: left; margin-right: 5px; }
        .top-news h3, .news h3 { font-size: large; font-weight: bold; }
        .accesskey { text-decoration: underline; }
        a { text-decoration: none; }
        .validation { margin-top: 10px; }

        .product img { float: left; margin-right: 5px; }
        .product h3, .news h3 { font-size: large; font-weight: bold; }
    </style>
    
    <% 
    Response.AddHeader "Refresh", "60" 
    %>

  </head>
  <body>
  
<%

	Dim rs
    Dim SQL
	Dim connBarleyReadOnly
	
	set connBarleyReadOnly=server.CreateObject("ADODB.Connection")
	connBarleyReadOnly.Mode = 1
 	connBarleyReadOnly.Open "driver=SQL Server;server=68.178.202.54;uid=davewj2o;pwd=get2it;database=ha_prod"	

    set rs = server.createObject("ADODB.Recordset")
	SQL = "ni_DailyActivity('" & FormatDateTime(Date(),2) & _
	"', '" & FormatDateTime(Date(),2) & "')"

	rs.open sql,connBarleyReadOnly,3,3,4

%>

    <div class="header">
      <img src="includes/images/logo_harvest_mobi.gif" alt="Harvest American, Inc." /><br />
      <small><%=rs("ipdate")%></small>
    </div>

    <div class="content">

        <table width="180" border="0" align="center">
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
        
   </div>
   
    <div class="footer">
       <small>
         &copy; <% Response.Write Year(now) %> Harvest American, Inc. <br />All rights reserved.<br />
       </small>
    </div>

    <!--<div class="validation">
      <a href="http://validator.w3.org/check?uri=referer">
        <img src="http://www.w3.org/Icons/valid-xhtml10" alt="Valid XHTML-MP" height="31" width="88" />
      </a>
      <a href="http://jigsaw.w3.org/css-validator/check/referer">
        <img src="http://jigsaw.w3.org/css-validator/images/vcss" alt="Valid CSS" width="88" height="31" />
      </a>
    </div>-->
  </body>
</html>