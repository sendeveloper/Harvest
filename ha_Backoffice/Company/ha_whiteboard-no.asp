<% Option Explicit %>
<!DOCTYPE html>
<!--#include virtual="ha_BackOffice/includes/sql.asp"-->
<% 
  Dim strPath, strPathBase, strPathBOBase, strPathIncludes, strPathBOIncludes, strPathImages, strPathServers, strPathTelephones, strPathNotifications, strPathEmployees, strPathCompany, strPathTableDistribution, PageHeading, referrer, strPathControlPanel, ColorTab, strPathOld, strPathBaseOld, strPathBOBaseOld, strPathIncludesOld, strPathBOIncludesOld, strPathImagesOld, strPathServersOld, strPathTelephonesOld, strPathNotificationsOld, strPathEmployeesOld, strPathCompanyOld
  RowMod = 3 
  Session("Redirect") = ""
  ColorTab = 5
  PageHeading = "Whiteboard"

  referrer = Request("referrer")
  If referrer = "" Then
    referrer = Request.Servervariables("HTTP_REFERER")
  End If
%>
<!--#include virtual="ha_BackOffice/includes/lib.asp"-->
<!--#include  virtual="ha_BackOffice/includes/config.asp"-->
<html>
  <head>
    <title>Harvest American BackOffice - <%=PageHeading%></title>

    <meta charset="UTF-8"></meta>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
    <meta name="description" content="Whiteboard of current priorities">

    <link href="<%=strPathBOIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet">
    <link href="<%=strPathControlPanel%>includes/HarvestAmerican.css" type="text/css" rel="stylesheet">
    <link rel="stylesheet" type="text/css" href="<%=strPathBOIncludes%>menu/dropdowntabfiles/halfmoontabs.css">

    <script type="text/javascript" src="<%=strPathBOIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>
    <script type="text/javascript" src="<%=strPathBOIncludes%>lib.js"></script>

    <style id="top">
      html {width: 100%;}
      body {white-space: nowrap;}
      header {}
      /* nav {height: 20%;} */
      aside {display: inline-block; width: 18%; overflow: auto; border-right: 1px solid blue; vertical-align: top;}
      article {display: inline-block; width: 100%; overflow: auto; height: 100%; margin-top: 0em;}
      footer {}
      div.wonky-table-look-alike {margin-left: auto; margin-right: auto; width: 78%; overflow: hidden; text-align: left; padding: 0em .5em .5em .5em}
      h {text-align: center; font-weight: bold; font-size: 14pt; display: inline-block;}
      aside > h {border-bottom: 1px solid black; margin-top: 1em;}
      aside > h:first-child {margin-top: 0em;}
      table.resultset {background: lightblue;}
      table.resultset tr.rowmod-0 {background: lightyellow;}
      table.resultset tr.rowmod-1 {background: lightpink;}
      table.resultset tr.rowmod-2 {background: lightgreen;}
      aside > select {display: block;}
  
      .done {text-decoration: strikethrough; background: darkgray;}

      /* Temporarily make it readable for development */
      /* div.wonky-table-look-alike {width: 100%;} */
  
      .layout {list-style: none;}
    </style>
    <script type="text/javascript" language="javascript">
      function init(){
        alert(get("table.resultset tr td")[0][1] //.forEach(function (e) {
          //e[5].innerHTML})
      );
        return;}
      window.onload = init;
    </script>
  </head>
  <body class="gray_desktop">
    <header></header>
    <div class="wonky-table-look-alike MainBody">
    <article>
      <nav><!--#include virtual="ha_Backoffice/includes/bodyparts/ha_PageHeading.inc"--></nav>
<%
  sqlConn.Close()
  'on error resume next
  sqlConn.Open("Driver=SQL Server;server=72.43.236.2,3433;uid=davewj2o;pwd=get2it;database=tempdb") 'camden3 
  'sqlTableInsert("declare @table table (Priority int, Description nvarchar(max), ""Assigned To"" nvarchar(max), Server nvarchar(max), Category nvarchar(max), Status nvarchar(max)); insert into @table select * from openrowset('msdasql','DRIVER=Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb);UserCommitSync=Yes;Threads=3;SafeTransactions=0;ReadOnly=1;PageTimeout=5;MaxScanRows=8;MaxBufferSize=2048;FIL=excel 12.0;DriverId=1046;colnameheader=false;DefaultDir=G:\office-share\Personal\Angel\Staff Meeting Minutes 2011\;DBQ=ToDoList_2013_07_11.xlsx', 'select * from [ToDoListBreakRoom$]'); delete from @table where Description = 'Description'; select * from @table") '-- ToDoListBreakRoom -- RateLookupToDoList
  sqlTableInsert("declare @table table (Priority int, Description nvarchar(max), ""Assigned To"" nvarchar(max), Server nvarchar(max), Category nvarchar(max), Status nvarchar(max)); insert into @table select * from (values ('1', 'Description', 'Assigned To', 'Server', 'Category', 'Status')) as l(a,b,c,d,e,f); select * from @table")'-- ToDoListBreakRoom -- RateLookupToDoList"
'  If Error Then 
'    sqlConn.Close()
'    Respones.Write("Error"): Response.End
'   End If
  sqlConn.Close()
'  on error goto 0
%>
    </article>
    </div>
    <aside></aside>
    <footer></footer>
  </body>
</html>

<!--  -->
