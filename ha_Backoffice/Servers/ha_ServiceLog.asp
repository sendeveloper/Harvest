<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

<!--#include virtual="includes/connection.asp"-->

<html>
<head>

    <TITLE>Harvest American Service Log</TITLE>
    <META http-equiv="Content-Type" content="text/html; charset=UTF-8">
    <META http-equiv="Cache-Control" content="no-cache">


    <script type="text/JavaScript" src="<%=strPathIncludes%>haFunctions.js"></script>

    <link href="<%=strPathControlPanel%>includes/HarvestAmerican.css" type="text/css" rel="stylesheet">


<style type="text/css">


td.Heading 
        {
        font-size: 14px;
        font-weight: bold;
	text-align: center;
        }

table.Properties
	{
	font-size: 8pt;
	border: 1px solid black;
	}

table.PropertiesSection
	{
	border: 1px solid gray;
	}

</style>

<script language="javascript" type = "text/javascript">


</script>

</head>


<body onLoad="SetScreen(450, 500);">

<table width="400" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr valign="top">
    <td>

      <table width='100%' border='0' cellspacing='2' cellpadding='2'>
        <tr>
          <td class='Heading'>
            <%=Request("Task")%>
          </td>
        </tr>
      </table>

<%
    set rs=Server.CreateObject("ADODB.Recordset")

    'Read Properties
    'sql = "ha_PingServer('" & Request("ServerID") & "')"
    'rs.open sql, objcon, 3, 3, 4

%>

      <table width='400' cellspacing='0' cellpadding='0' class="Properties">
        <tr>
          <td>

        <%
'          if not RS.EOF then
'            do while not rs.eof
        %>

              <br>

        <%
'              rs.MoveNext
'            loop
'          end if
'          rs.close
        %>

          </td>
        </tr>
      </table>

      <table width="100%" border="0" cellspacing="2" cellpadding="2">
        <tr>
          <td width="50%">&nbsp;</td>
          <td>
            <br><br>
            <a href="javascript:window.close();" class="button">Close</a>
          </td>
          <td width="50%">&nbsp;</td>
        </td>

  </tr>
  <tr>
    <td>
      <div style="margin-left: 5em; height: 4em"><!--#include virtual="ha_BackOffice/includes/BodyParts/copyright.asp"--></div>
    </td>
  </tr>
</table>

</body>
</html>
