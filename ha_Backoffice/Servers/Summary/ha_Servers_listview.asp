<%
  Session("Redirect") = ""
  ColorTab = 1
   PageHeading = "Servers Summary | Listview"
%>
<!--#include virtual="/includes/connection.asp"--> <!-- config.asp overrides the paths set in connection.asp -->
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

  <link href="<%=strPathIncludes%>HarvestAmerican.css" type="text/css" rel="stylesheet">
  <link href="<%=strPathIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet">
  <link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>menu/dropdowntabfiles/halfmoontabs.css">

  <script type="text/javascript" src="<%=strPathIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>
  <script type="text/javascript" src="<%=strPathBOIncludes%>haFunctions.js" language="javascript"></script>
  <script language="javascript" type = "text/javascript">
    
    function status(ip) {
      openPopUp("server.asp");}
  </script>
  <style type="text/css">
    th {text-align: left; border-bottom: 1px solid black;}
    th.Heading1 {font-size: 14px; text-align: center; border-bottom: 1px solid black;}

    div.resultset {margin-top: 0.50em; margin-bottom: 4.0em; margin-left: 2em; margin-right: 2em; border-bottom: 1px solid black;}
    table.resultset {margin-bottom: 1.75em; margin-left: auto; margin-right: auto; width: 100%; border-spacing: .3em;}
    table.resultset td, table.resultset th {padding-top: 0.50em; padding-bottom: 0.51em;}
    table.resultset th.source, table.resultset th.last-run {text-align: center; font-size: 90%; padding-top: 0.8em; padding-bottom: .25em;}

    td.id, th.id,
    td.TaskLogName, th.TaskLogName,
    td.DurationIncrement, th.DurationIncrement,
    td.LastRunDateTime, th.LastRunDateTime {display: none;}

    th {width: 10%;}
    th.TableLocation, th.LastRunDuration, th.ActionButton {width: 20%;}
    th.ActionButtons {width: auto;}

    .views {width: 100%; text-align: right; font-size: 80%;}
    .views a {margin-right: 2em;}
  </style>
</head>
<body class="gray_desktop">
  <div class="MainBody">
    <!--#include virtual="ha_Backoffice/includes/bodyparts/ha_PageHeading.inc"-->
    <div class="views"><a href="ha_Servers.asp">Box View</a></div>
<%
    ' Collect server properties into rows from ha_Types table
    Dim vbcrlf: vbcrlf = chr(13) & chr(10)
    Dim rs: Set rs = Sql(_
      "select Description=min(Description) " &_
      "from ha_Backoffice.dbo.ha_Types_repl " &_
      "where Class like '% - Properties' " &_
      "group by Description order by min(Sequence)")
    Dim JoinProperties: JoinProperties = ""
    Dim JoinSelection: JoinSelection = ""
    Do While Not rs.eof
      JoinSelection = JoinSelection & ", [" & rs("Description") & "]=coalesce([" & rs("Description") & "].Value, '---')"
      JoinProperties = JoinProperties &_
        "left join ha_Backoffice.dbo.ha_Types_repl as [" & rs("Description") & "] " &_ 
        "on [" & rs("Description") & "].Class = server.Class " &_
        "and [" & rs("Description") & "].Description = '" & rs("Description") & "' "
      rs.MoveNext
    Loop

    sqlTableInsert(_
       "select distinct server=replace(server.Class, ' - Properties', '')" & JoinSelection & ", " &_
       "Action='<a class=""button"" href=""javascript:void(status(''' + [IP].Value + '''))"">Status<a/>' " &_
       "from (select Class from ha_Backoffice.dbo.ha_Types_repl) as server " &_
       joinProperties &_
       "where server.Class like '% - Properties' " &_
       "order by replace(server.Class, ' - Properties', '')")
%>
  </div><!-- MainBody -->
  
  <!--#include virtual="ha_BackOffice/includes/BodyParts/ha_PageTail.inc"-->
</body>
</html>
