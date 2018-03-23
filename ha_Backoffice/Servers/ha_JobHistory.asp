<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

<!--#include virtual="includes/connection.asp"-->

<html>
<head>

    <TITLE>Harvest American Server Job History</TITLE>
    <META http-equiv="Content-Type" content="text/html; charset=UTF-8">
    <META http-equiv="Cache-Control" content="no-cache">


    <script language="JavaScript" src="<%=strPathIncludes%>haFunctions.js"></script>

    <link href="<%=strPathControlPanel%>includes/HarvestAmerican.css" type="text/css" rel="stylesheet">


<style type="text/css">


td.Heading 
        {
        font-size: 14px;
        font-weight: bold;
	text-align: center;
        }

table.PropertiesSection
	{
	border: 1px solid gray;
	}

</style>

<script language="javascript" type = "text/javascript">


</script>

</head>


<body onLoad="SetScreen(650, 700);">

<table width="400" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr valign="top">
    <td>

      <table width='100%' border='0' cellspacing='2' cellpadding='2'>
        <tr>
          <td class='Heading'>
            Job History
          </td>
        </tr>
      </table>

<%
    set rs=Server.CreateObject("ADODB.Recordset")

    if Request("job_id") > "" then
        Set conn=Server.CreateObject("ADODB.Connection")

        Select Case Request("Server")
        Case "Barley1"
            conn.Open "driver=SQL Server;server=Barley1.HarvestAmerican.net;uid=davewj2o;pwd=get2it;database=Harvest"
        Case "Barley2"
            conn.Open "driver=SQL Server;server=Barley2.HarvestAmerican.net;uid=davewj2o;pwd=get2it;database=Harvest"
        Case "Camden1"
            conn.Open "driver=SQL Server;server=Camden1.HarvestAmerican.net;uid=davewj2o;pwd=get2it;database=SSIS_Logs"
        End Select

        'Read Properties
        SQL = "ha_ControlPanel_Job_properties('" & Request("job_id") & "')"
        rs.open SQL, conn, 3, 3, 4
        if not rs.EOF then
            jName = rs("Name")
            jDescription = rs("Description")
        End If
        rs.close
    end if

%>

      <table width='600' border="0" cellspacing='2' cellpadding='2'>
        <tr valign="top">
          <td width="20%" align="right">
            Server:
          </td>
          <td width="80%">
            <%=Request("Server")%>
          </td>
        </tr>
        <tr valign="top">
          <td width="20%" align="right">
            Job Name:
          </td>
          <td width="80%">
            <%=jName%>
          </td>
        </tr>
        <tr valign="top">
          <td width="20%" align="right">
            Description:
          </td>
          <td width="80%">
            <%=jDescription%>
            <br><br>
          </td>
        </tr>
      </table>

      <table width='600' cellspacing='0' cellpadding='0' style="border: 1px solid black;">
        <tr>
          <td>
            <table width='100%' cellspacing='0' cellpadding='2'>

<%
    If jName > "" Then
        SQL = "ha_ControlPanel_Job_history('" & Request("job_id") & "')"
        rs.open SQL, conn, 3, 3, 4

        if not rs.EOF then
            do while not rs.eof
%>

              <tr valign="top">
                <td width="20%">
                  <%=rs("run_date_time")%>
                </td>
                <td width="80%">
                  <%=rs("message")%>
                </td>
              </tr>

<%
                rs.MoveNext
            loop
        end if
        rs.close
    End If
%>

            </table>
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
      </table>

    </td>
  </tr>
  
</table>
<!--#include virtual="ha_BackOffice/includes/BodyParts/ha_PageTail.inc"-->
</body>
</html>
