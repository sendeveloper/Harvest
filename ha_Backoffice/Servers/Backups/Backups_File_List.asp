<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

<!--#include virtual="includes/connection.asp"-->
<!--#include virtual="ha_backoffice/includes/Config.asp"-->

<%
	dServer = Request("s")
	dDatabaseName = Request("n")
	dDir = Request("d")
	dDir = Replace(dDir,"*","\")
%>

<html>
<head>
    <title>Server Backups File List</title>

    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>ha_BackofficePopup.css">

    <script language="javascript" type="text/javascript" src="<%=strPathBOIncludes%>haFunctions.js"></script>
		
<script language='JavaScript'>
		
    function formLoad()
        {
        SetScreen(550, 600);
		}
				
</script>

</head>

<body onLoad="formLoad();">

<table width="500" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td>

        <table width="100%" border="0" cellspacing="5" cellpadding="5" align="center">

          <tr>
            <td class="popupHeading">
              Server Backups File List
            </td>
          </tr>
		
          <tr>
		    <td class="databaseHead" colspan="3">
			  <%=dDatabaseName%>
			  <span style="font-size: 8pt; color: #C0C0C0; margin: 4px 0 2px 0;"><br>(<%=dServer%> - <%=dDir%>)</span>
			</td>
		  </tr>

		  <tr>
		    <td>
				<div style="height:380px;
					vertical-align: top;
					text-align: left;
					overflow-y: scroll;">
                  <table cellspacing="0" cellpadding="2" id="sectionLines">


<%
	set rs=Server.CreateObject("ADODB.Recordset")
    sql = "ha_Backup_File_list('" & dDatabaseName & "', '" & dServer & "', '" & dDir & "')"
    rs.open sql,ConnCasper10, 3, 3, 4

    If not rs.eof then
		LineCount = 0
		do while not rs.eof
			LineCount = LineCount + 1
%>
		<tr valign="top">
		  <td width="8%">
            <%=LineCount%>
		  </td>
		  <td Width"92%">
            <%=rs("subdirectory")%>
		  </td>
		</tr>

<%
			rs.movenext
		loop
    End If	    

    rs.close
%>

                  </table>
				</div>
		  
			</td>
		  </tr>

        </table>	
		
 
        <table width="100%" border="0" cellspacing="5" cellpadding="5">
		  <tr>
		    <td width="45%">
			  &nbsp;
			</td>
		    <td align="center">
                <a href="javascript:close();" class="bo_Button100" title="Closes this window">Ok</a>
		    </td>
		    <td width="45%">
			  &nbsp;
			</td>
		  </tr>
        </table>
		
    </td>
  </tr>
</table>

</body>
</html>
