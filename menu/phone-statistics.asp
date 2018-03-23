
<html>

<!--#include virtual="includes/connection.asp"-->
<!--#include virtual="menu/includes/phone-functions.asp"-->

<head>
    <title>Harvest American Phone Call Log Statistics</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <script language="JavaScript" src="includes/checkDate.js" type="text/javascript"></script>
    <script language="JavaScript" src="datepick/ts_picker.js" type="text/javascript"></script>
    <script language="JavaScript" src="<%=strPathIncludes%>haFunctions.js"></script>

    <link href="<%=strPathControlPanel%>includes/HarvestAmerican.css" type="text/css" rel="stylesheet">

<%
'    Response.buffer=true

    Dim referrer
    referrer = Request("referrer")
    If referrer = "" Then
      referrer = Request.Servervariables("http_referrer")
    End If


   Dim frequency
   frequency = Request("freq")
   If frequency = "" Then frequency = "Day"
%>
</head>

<body>
<%=referrer%>
<table width="100%" align="center" border="0" cellspacing="0" cellpadding="0">

  <tr> 
    <td width="80%" align="left"> 
      <a href="phone-statistics.asp?freq=<%=Iif(frequency = "Day", "Week", "Day")%>"><%=Iif(frequency = "Day", "Week", "Day")%></a>
      <h1>Harvest American Phone Call Log Statistics</h1>
    </td>
    <td width="20%" align="right">
      <a href="phone-statistics.asp" class="button">Update</a>
      <a href="phones.asp" class="button">Back</a>
    </td>
  </tr>

  <tr>
    <td colspan="2">
      <table class="table">
<%
  Dim recordCount
  recordCount = 0

  objcon.Execute("ha_Phone_Statistics_updateRecent")
  objcon.Close
  set objcon=server.CreateObject("ADODB.Connection")
  objcon.Open "driver=SQL Server;server=dallas01.harvestamerican.net;uid=davewj2o;pwd=get2it;database=ha_BackOffice"

  Dim rs
  Set rs = server.createObject("ADODB.Recordset")
  rs.open "ha_Phone_Statistics_view('" & frequency & "')", objcon, 3, 3, 4
  
  Do While Not rs.eof
%>
        <tr>
          <td><%=rs("statisticdate")%></td>
          <%=rs("statisticcalls")%>
        </tr>
<%
   rs.movenext
   Loop
%>
      </table>
    </td>
  </tr>
</table>
</body>
</html>

<%
Function IIf(condition, consequent, alternative)                   
  If condition Then IIf = consequent Else IIf = alternative
End Function

%>
