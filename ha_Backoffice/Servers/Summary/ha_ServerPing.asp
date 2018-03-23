<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_Backoffice/includes/config.asp"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->
  
<%
	Title = "Server Ping"	
%>

<html>
<head>
  <title>Harvest American <%=Title%></title>
  <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
  <meta http-equiv="Cache-Control" content="no-cache">
  
  <link href="/ha_Backoffice/includes/ha_BackOfficePopup.css" type="text/css" rel="stylesheet">
  <script language="javascript" type="text/javascript" src="/includes/haFunctions.js"></script>
  
  <style type="text/css">
    td.Heading 
        {
        font-size: 12px;
        font-weight: bold;
        }

    table.Properties
		{
		font-size: 8pt;
		border: 1px solid gray;
		margin: 0 auto;
		}
  </style>
  
</head>


<body onLoad="SetScreen(450, 500);">

	  <span class="popupHeading"><%=Title%></span>

  <table width="95%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr valign="top">
      <td>
	
		<table width='100%' border='0' cellspacing='2' cellpadding='2'>
	      <tr><td>&nbsp;</td></tr>
          <tr>
            <td class='Heading'>
              Ping from Casper10 to <%=Request("ServerID")%>
            </td>
          </tr>
		</table>
	
<%
    set rs=Server.CreateObject("ADODB.Recordset")

    'Read Properties
    sql = "ha_PingServer('" & Request("ServerID") & "')"
    rs.open sql, connCasper10, 3, 3, 4

%>

        <table width='100%' cellspacing='0' cellpadding='0' class="Properties">
          <tr>
            <td style="padding-left: 5px;">
        <%
          if not RS.EOF then
            do while not rs.eof
        %>
              <%=rs("Output")%><br>
        <%
              rs.MoveNext
            loop
          end if
          rs.close
        %>
            </td>
          </tr>
	    </table>

      </td>
    </tr>
    <tr><td>&nbsp;</td></tr>
  </table>
  
    <span class="popupButtons" style="padding-bottom: 20px;">
      <a href="javascript:window.close();" class="bo_Button80">Close</a>
    </span>
  
</body>
</html>
