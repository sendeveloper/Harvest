<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

<!--#include virtual="includes/connection.asp"-->
<!--#include virtual="ha_Backoffice/includes/config.asp"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->
  
<%
	Title = "Server Properties"	

	p = instr(Request("ServerID"),"-")
	if p then
		ServerName = left(Request("ServerID"),p-2)
	else
		ServerName = Request("ServerID")
	end if				
	
%>

<html>
<head>

    <title>Harvest American <%=Title%></title>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
    <meta http-equiv="Cache-Control" content="no-cache" />


    <script type="text/javaScript" src="<%=strPathIncludes%>haFunctions.js"></script>

    <link href="<%=strPathControlPanel%>includes/HarvestAmerican.css" type="text/css" rel="stylesheet" />
    <link href="<%=strPathBOIncludes%>ha_Backoffice.css" type="text/css" rel="Stylesheet" />

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

    table.PropertiesSection
		{
		border: 1px solid #DDDDDD;
		}
  </style>
  
</head>

<body onload="SetScreen(500, 675);">

	  <span class="popupHeading"><%=Title%> - <%=ServerName%></span>

<table width="95%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr><td>&nbsp;</td></tr>
  <tr valign="top">
    <td>

    <%
    set rs=Server.CreateObject("ADODB.Recordset")

    'Read Properties
    sql = "SELECT * from ha_Types WHERE Class = '" & Request("ServerID") & "' ORDER BY Sequence"
    rs.open sql, connPubTables, 1, 2

    IsSQLServer = "No"
    DomainName = ""
	SQLServerPort = ""
    %>

      <table width='100%' cellspacing='0' cellpadding='0' class="Properties">
        <tr>
          <td>

        <%
          if not RS.EOF then
            do while not rs.eof
        %>

            <table width='100%' cellspacing='0' cellpadding='2' class="PropertiesSection">
              <tr valign="top">
                <td width="50%">
                  <%=rs("Description")%>
                </td>
                <td width="50%">
                  <%=rs("Value")%>
                </td>
              </tr>
            </table>

        <%
              if rs("Description") = "SQL Server" and rs("Value") = "Yes" then
                IsSQLServer = "Yes"
              end if
              if rs("Description") = "Domain Name" then
                DomainName = rs("Value")
              end if
              if rs("Description") = "SQL Server Port" then
                SQLServerPort = rs("Value")
              end if
			  

              rs.MoveNext
            loop
          end if
          rs.close
        %>

          </td>
        </tr>
      </table>


<%
    If IsSQLServer = "Yes" and DomainName > "" then

      'Create heading
      p = instr(Request("ServerID"),"-")
      if p then
        sqlHead = left(Request("ServerID"),p+1) & "SQL Server"
      else
        sqlHead = Request("ServerID")
      end if
	  
		If ServerName <> "Philly02" and ServerName <> "Philly03" and ServerName <> "Philly04" Then

		  Dim objConn: Set objConn = Server.CreateObject("ADODB.Connection")
		  Dim strConn: strConn = "driver=SQL Server;server=" & DomainName & IIf(SQLServerPort > "", ", " & SQLServerPort, "") & ";uid=davewj2o;pwd=get2it;database=master"
		  'Response.write strConn & "<br>"
		  objConn.Open strConn
	  
          'Read Properties
          sql = "select SERVERPROPERTY('Edition') as [Edition], " & _
            "SERVERPROPERTY('ProductVersion') as [Product Version], " & _
            "SERVERPROPERTY('ProductLevel') as [Product Level]"
          rs.open sql, objConn, 1, 2

%>
      <table width='100%' cellspacing='0' cellpadding='0'>
        <tr>
          <td class='Heading'>
            <br>
            <%=sqlHead%>
          </td>
        </tr>
      </table>

      <table width='100%' cellspacing='0' cellpadding='0' class="Properties">
        <tr>
          <td>

<%
              For i = 0 to rs.fields.count-1
%>
            <table width='100%' cellspacing='0' cellpadding='2' class="PropertiesSection">
              <tr valign="top">
                <td width="50%">
                  <%=rs.fields(i).name%>
                </td>
                <td width="50%">
                  <%=rs.fields(i).value%>
                </td>
              </tr>
            </table>
<%
              Next
              rs.close
              objConn.close
		End If
%>
          </td>
        </tr>
      </table>

<%
    End If
%>

    </td>
  </tr>
  <tr><td>&nbsp;</td></tr>
</table>

    <span class="popupButtons" style="padding-bottom: 20px;">
      <a href="javascript:window.close();" class="bo_Button80">Close</a>
    </span>

</body>
</html>
