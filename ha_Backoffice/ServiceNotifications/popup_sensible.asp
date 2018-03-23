<%
  Session("Redirect") = ""
  PageHeading = "Service Notification Edit"
%>
<!--#include virtual="/includes/connection.asp"-->
<!--#include virtual="/ha_backoffice/includes/Config.asp"-->
<!--#include virtual="/ha_backoffice/includes/sql.asp"-->
<%
  rowmod = 2
  ShowRowNum = False
  ShowRowCount = False
%>

<html>
<head>
  <title>Harvest American Backoffice - <%=PageHeading%></title>
  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
  <link href="<%=strPathIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet">
  <link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>menu/dropdowntabfiles/halfmoontabs.css">
  <script type="text/javascript" src="<%=strPathIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>
  <style type="text/css">
    header {display: block; text-align: center; width: 100%; margin: 0em; padding: 0em;}
    header > h1 {display: inline-block; text-align: center; border-bottom: 2px solid black; padding-bottom: 0.10em; font-size: 110%; width: 90%}
    article {display: block; text-align: center; width: 100%; margin: 0em; padding: 0em;}
    article > div {display: inline-block; width: 90%; margin: 0em; padding: 0em;}
    footer > ul {display: block; text-align: center;}
    footer > ul > li {display: inline-block; margin-left: 2em; margin-right: 2em; width: 5em;}

    .undo, .redo {visibility: hidden;}

    .resultset tfoot th {border: none;}

    th {text-align: left; border-bottom: 1px solid black; color: lightgray;}
    th.Heading1 {font-size: 14px; text-align: center; border-bottom: 1px solid black; color: black;}

    div.resultset {margin-top: -.5em; border-bottom: 1px solid black; width: 100%;}
    table.resultset {margin-top: 0em; margin-bottom: 1.75em; margin-left: auto; margin-right: auto; width: 100%; border-spacing: .3em;}
    table.resultset > thead > tr:first-child {border-spacing: 0em;}
    table.resultset td, table.resultset th {padding-top: 0.50em; padding-bottom: 0.51em;}
    table.resultset th.service, table.resultset th.notifee-heading, table.resultset th.service-name {text-align: center; font-size: 90%; padding-top: 0.8em; padding-bottom: .25em; color: black;}

    td.id, th.id,
    td.TaskLogName, th.TaskLogName,
    td.DurationIncrement, th.DurationIncrement,
    td.LastRunDateTime, th.LastRunDateTime {display: none;}

    th {width: 10%; height: auto;}
    th.TableLocation, th.LastRunDuration, th.ActionButton {width: 20%;}
    th.ActionButtons {width: auto;}
  </style>
</head>
<body>
  <header>
    <h1><%=PageHeading%></h1>
  </header>
  <article>
    <div>
<%
  If Request("service") = "" Then
%>
      <form>
        <label for="server">Server</label>
        <input id="server" type="text">
        <label for="service">Service Name</label>
        <input id="service" type="text">
      </form>
<%
  End If
%>

<%
' Tediously create a renaming map for the columns because ASP is horribly broken.
Dim columnSpec: columnspec = Split("Table=TableName, Location=TableLocation, Frequency, Date=LastRunDate, Time=LastRunTime, Duration Since Last=LastRunDuration, Action=ActionButton", ",")
Dim column: Set column = Server.CreateObject("Scripting.Dictionary")
For Each col in columnspec
  Dim renaming: renaming = Split(Trim(col), "=")
  If Ubound(renaming) > 0 Then
    Call column.Add(renaming(1), renaming(0))
  Else
    Call column.Add(renaming(0), renaming(0))
  End If
Next

Function renameColumns (field)
  renameColumns = iif(column.Exists(field), column(field), field)
End Function

Function iif(condition, consequent, alternative)
  If condition Then iif = consequent Else iif = alternative
End Function

   Dim rs: Set rs = sql("select id = sequence + '-' + colid, LastCheckPassed, [Max Delay (s)]=MaxDelay, Notifee, Sequence, Priority, Status = case Enabled when 1 then 'Enabled' else 'Disabled' end  from casper10.ha_BackOffice.dbo.ha_service_NotificationView where Service = '" & CStr(Request("service")) & "' order by service, sequence, colid")
   'Set sqlHeaderDecorator = GetRef("renameColumns")
%>
    <%=replace(rsTable(rs), "<thead>", "<thead><tr><th class=""service"">For Service: </th><th class=""service-name"" colspan=""2"">" & CStr(Request("service")) & "</th><th class=""notifee-heading"" colspan=""3"">Notifee</th><th></th></tr>" )%>
    <!--
    <tr>
      <th class="Heading1" colspan="2">Source</th>
      <th class="Heading1">&nbsp;</th>
      <th class="Heading1" colspan="3">Last Run</th>
      <th class="Heading1">&nbsp;</th>
    </tr>

          <th width="10%">Table</th>
          <th width="20%">Location</th>
          <th width="10%">Frequency</th>
          <th width="10%">Date</th>
          <th width="10%">Time</th>
          <th width="20%">Duration Since Last</th>
          <th width="20%">Action</th>
     -->
    </div>
  </article>
  <footer>
    <ul>
      <li><button class="save" disabled="disabled">Save</button></li>
      <li><button class="undo" disabled="disabled">Undo</button></li>
      <li><button class="redo" disabled="disabled">Redo</button></li>
      <li><button class="revert">Revert</button></li>
      <li><button class="cancel">Close</li>
    </ul>
  </footer>
</body>
</html>

<script type="text/javascript" src="<%=strPathBOIncludes%>haFunctions.js" language="javascript"></script>
<script type="text/javascript" src="<%=strPathIncludes%>lib.js"></script>
<script type="text/javascript">
  node = nodey;

  listen(window, "load", function init(e) {
    listen(get(".save")[0][0], "click", function save(e) {
      get(".cancel")[0][0].innerHTML = "Close";
      get(".save")[0][0].disabled = true;
      get(".revert")[0][0].disabled = false;});
    listen(get(".revert")[0][0], "click", function revert(e) {
      get(".cancel")[0][0].innerHTML = "Cancel";
      get(".save")[0][0].disabled = false;
      get(".revert")[0][0].disabled = true;});
    listen(get(".cancel")[0][0], "click", function cancel(e) {window.close();});
  
    if (get("table.resultset tfoot")[0].length === 0) {
      get("table.resultset")[0][0].appendChild(node("tfoot"));}
    if (get("table.resultset tfoot tr")[0].length === 0) {
      get("table.resultset tfoot")[0][0].appendChild(node("tr"));}
    listen(get(".resultset tfoot tr")[0][0].appendChild(node("th", {}, node("button", {}, text("Add")))),
      "click", 
      function(e){
        get(".resultset tbody")[0][0].appendChild(node("tr", {}, ["name", "lastcheckpassed", "maxdelay", "notifee", "-", "priority", "status"].forEach(function(e){return node("td", {}, text("<new>"))})));});

    SetScreen(800, 800);});
</script>
