<%
  Session("Redirect") = ""
  ColorTab = 3
   PageHeading = "Service Alert Panel"
   %>
<!--#include virtual="ha_BackOffice/ServiceNotifications/top.inc"-->
<!--#include virtual="ha_backoffice/includes/Config.asp"-->
  <style type="text/css">
    td {
      font-family: Arial, Helvetica, sans-serif;
      font-size: 12px;
      font-style: normal;
      line-height: normal;
      font-weight: normal;
      font-variant: normal;
      text-transform: none;
      color: #000000;
      text-decoration: none;}

    table.ServerFarm {
      border: 2px solid black;
      background-color: #FFFFCC;
      width: 100%;
      height: 300px;}

    table.Server {
      border: 5px inset black;
      background-color: #DDDDDD;
      width: 100%;}

    td.ServerName {
      font-size: 14px;
      font-weight: bold;}

    tr.ServerType {
      font-size: 12px;
      text-align: center;}

    tr.ServerIP{
      font-size: 12px;
      text-align: center;}

    img.status {width: 3em; height: 3em;}
    .server {margin-bottom: 2em;}
    .server .layout ul {display: inline-block;}
    .server .layout li {display: inline-block;}
    .server-icon {display: inline-block;}
  </style>

  <script language="javascript" type = "text/javascript">
    function clickProperties(ServerID) {
      var URL = '<%=strPathControlPanel%>cp_ServerProperties.asp' +
                '?ServerID=' + ServerID;
      openPopUp(URL);}

    function clickPing(ServerID) {
      var URL = '<%=strPathControlPanel%>cp_ServerPing.asp' +
                '?ServerID=' + ServerID;
      openPopUp(URL);}
  </script>


<ul class="layout">
<%
  Dim rsFarm: Set rsFarm = Sql("select Farm = 'Zip2Tax'")
  Do While Not rsFarm.eof
%>

  <li>
    <ul class="layout">
<%
    Dim rsServer
    Set rsServer = Sql("select * from casper10.ha_BackOffice.dbo.ha_service_Notifications")
    Do While Not rsServer.eof
%>
      <li>
        <div class="server" style="display: inline-block;">
          <ul class="layout">
          <li class="graphic" style="display: inline-block;">
            <img class="server-icon" src="/ha_BackOffice/images/server.png" alt="server" />
          <li class="status"><img class="status" src="/ha_BackOffice/images/<%=iif(datediff("s", rsServer("LastCheckPassed"), Now()) < rsServer("MaxDelay"), "status-pass.png", "status-fail.png")%>"/><span class="timestamp"><%=rsServer("LastCheckPassed")%></span></li>
          <li class="message" style="display: <%=iif(DateDiff("m", rsServer("LastCheckPassed"), Now()) > 10, "none", "inline")%>"></li>
          <li class="details"><button onclick="alert('<%=rsServer("LastCheckPassed")%>');">Details</button></li></ul>
          <%=rsServer("Description")%>
        </div>
      </li>
<%
      rsServer.MoveNext
    Loop
%>
    </ul>
  </li>
<%
    rsFarm.MoveNext
  Loop
%>
</ul>

<!--#include virtual="ha_BackOffice/ServiceNotifications/bottom.inc"-->
