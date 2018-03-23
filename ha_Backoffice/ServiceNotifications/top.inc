<!DOCTYPE html>
<!--#include virtual="ha_BackOffice/includes/sql.asp"-->
<% 
  RowMod = 3 
%>
<!--#include virtual="ha_BackOffice/includes/lib.asp"-->

<html>
<head>
  <title>Harvest American BackOffice - <%=PageHeading%></title>
  <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<%
  referrer = Request("referrer")
  If referrer = "" Then
  referrer = Request.Servervariables("HTTP_REFERER")
  End If
%>
  <link href="<%=strPathBOIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet">
  <link href="<%=strPathControlPanel%>includes/HarvestAmerican.css" type="text/css" rel="stylesheet">
  <link rel="stylesheet" type="text/css" href="<%=strPathBOIncludes%>menu/dropdowntabfiles/halfmoontabs.css">
  <script type="text/javascript" src="<%=strPathBOIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>
  <script type="text/javascript" src="<%=strPathBOIncludes%>lib.js"></script>
  

  <style id="top">
    body {white-space: nowrap;}
    header {}
    /* nav {height: 20%;} */
    aside {display: inline-block; width: 18%; overflow: auto; border-right: 1px solid blue; vertical-align: top;}
    article {display: inline-block; width: 78%; overflow: auto; height: 100%; margin-top: 0em;}
    footer {}
    div.wonky-table-look-alike {margin-left: auto; margin-right: auto; overflow: hidden; text-align: left; padding: 0em .5em .5em .5em}
    h {text-align: center; font-weight: bold; font-size: 14pt; display: inline-block;}
    aside > h {border-bottom: 1px solid black; margin-top: 1em;}
    aside > h:first-child {margin-top: 0em;}
    table.resultset {background: lightblue;}
    table.resultset tr.rowmod-0 {background: lightyellow;}
    table.resultset tr.rowmod-1 {background: lightpink;}
    table.resultset tr.rowmod-2 {background: lightgreen;}
    aside > select {display: block;}

    /* Temporarily make it readable for development */
    /* div.wonky-table-look-alike {width: 100%;} */

    .layout {list-style: none;}
  </style>

  <script type="text/javascript" language="javascript">
    function init(){
      return;
    }
    window.onload = init;
  </script>
</head>

<body class="gray_desktop">
  
  <header></header>
<div class="wonky-table-look-alike MainBody">
  <nav><!--#include virtual="ha_Backoffice/includes/bodyparts/ha_PageHeading.inc"--></nav>

  <aside>
<%
  Dim rs
  Set rs = sql("select * from casper10.ha_BackOffice.dbo.ha_service_Notifications")
%>
    <h>Services</h>
    <select>
<%
   Do While not rs.eof
%>
      <option value="<%=rs("Description")%>"><%=rs("Description")%></option>
<%
     rs.MoveNext
   Loop
%>
    </select>
    <h>Notifees</h>
    <select>
<%
  Set rs = sql("select name = e.firstname, method = m.Description from casper10.ha_BackOffice.dbo.ha_service_Notifees as n join casper10.ha_BackOffice.dbo.ha_EmpAccounts as e on e.EmpId = n.EmpId join casper10.ha_BackOffice.dbo.ha_service_NotificationMethods as m on m.id = n.NotificationMethod")
  Do While Not rs.eof
%>
      <option>
        <%=rs("name")%> via <%=rs("method")%>
      </option>
<%
    rs.MoveNext
  Loop
  rs.Close
  Set rs = sql("select method = Description from casper10.ha_BackOffice.dbo.ha_service_NotificationMethods")
%>
    </select>
    <h>Methods</h>
    <select>
<% Do While Not rs.eof %>
      <option>
        <%=rs("method")%>
      </option>
<% 
     rs.MoveNext 
   Loop
   rs.Close
%>

    </select>
  </aside>

  <article>
