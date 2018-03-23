<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

<html>
<head>
  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/Config.asp"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->

<%
	Session("Redirect") = ""
	ColorTab = 1
	PageHeading = "Server Backups (Database)"

	If Request("Server") = "" or isNull(Request("Server")) Then
		WorkingServer = "Casper10"
	Else
		WorkingServer = Request("Server")
	End If
%>

  <title>Harvest American Backoffice - <%=PageHeading%></title>
  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

  <link href="<%=strPathIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet">
  <link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>menu/dropdowntabfiles/halfmoontabs.css">

  <script type="text/javascript" src="<%=strPathIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>
  <script type="text/javascript" src="<%=strPathBOIncludes%>haFunctions.js" language="javascript"></script>
  
  <script type="text/javascript">
  
	function changeServer()
		{
		var s = document.getElementById('SelectedServer').value;
		var URL = 'ha_Servers_Backups_Database.asp' +
			'?Server=' + s;
		window.location = URL;
		}
		
    function clickDatabaseBackupDiagram()
        {
        var URL = 'ha_Servers_Backups_Database_diagram.asp'
        openPopUp(URL);
        }
		
    function clickDatabaseBackupProcess()
        {
        var URL = 'ha_Servers_Backups_Database_process.asp'
        openPopUp(URL);
        }
		
    function clickDatabaseBackupPurgeDiagram()
        {
        var URL = 'ha_Servers_Backups_Database_Purge_diagram.asp'
        openPopUp(URL);
        }
		
    function clickPopup(databaseName, bServer, bDir)
        {
        var URL = 'Backups/Backups_File_List.asp' +
			'?n=' + databaseName +
			'&s=' + bServer +
			'&d=' + bDir;
		//alert(databaseName + ' | ' + bServer + ' | ' + bDir);
        openPopUp(URL);
        }

  </script>
  
  <style type="text/css">
    th
		{
		text-align:		center;
		border-bottom:	1px solid black;
		}

    .centered {margin-left: auto; margin-right: auto;}
    .centering {text-align: center;}

    body {height: 100%;}
    div.MainBody {height: 95%;}
  </style>
</head>

<body class="gray_desktop">
  <div class="MainBody">
  <!--#include virtual="ha_Backoffice/includes/bodyparts/ha_PageHeading.inc"-->
    <table width="95%" border="0" cellspacing="5" cellpadding="5" class="centered hybrid">

		
		<%=TopRightLinks(10, "", "Database Backup Process", "Database Backup Purge Diagram", "Database Backup Diagram")%>
        

	  <tr><td>&nbsp;</td></tr>
	  
	</table>
	  
	<table width="95%" border="0" cellspacing="5" cellpadding="5" class="centered hybrid">
			

			  <tr vAlign="top">
				<th width="18%">Created Date</th>
				<th width="12%">Results Status</th>
				<th width="12%">Elapsed Time</th>
				<th width="58%">Task</th>
			  </tr>

	<%
		recordCount = 0

		set rs=server.createObject("ADODB.Recordset")
		SQL = "ha_TaskLog_Backup_list"
		rs.Open SQL, connCasper10, 3, 3, 4

		if not RS.EOF then
			do while not rs.eof
				recordCount = recordCount + 1
				
				'Dividing Separator
				If SeparatorRows = "" Then
					'Defaults
					SeparatorRows = 3
					SeparatorWidth = "3px"
					SeparatorColor = "#D8F0E3"
				End If

				If recordCount mod SeparatorRows = 0 Then
					detailStyle = "border-bottom: " + SeparatorWidth + " solid " +SeparatorColor + ";"
				Else
					detailStyle = "border: 0;"
				End If
	%>
			
			  <tr>
				<td style="<%=detailStyle%>">
				  <%=rs("CreatedDate")%>
				</td>
				<td style="<%=detailStyle%>">
				  <%=rs("ResultStatus")%>
				</td>
				<td style="<%=detailStyle%>">
				  <%=rs("ElapsedTime")%>
				</td>
				<td style="<%=detailStyle%>">
				  <%=rs("Task")%>
				</td>
			  </tr>
			  
	<%
				rs.MoveNext
			loop
		end if
		
		rs.close
		connCasper10.close
	%>

	  <tr><td>&nbsp;</td></tr>			  
	</table>
			  
			  
  </div>
  <!--#include virtual="ha_BackOffice/includes/BodyParts/ha_PageTail.inc"-->
</body>
</html>

