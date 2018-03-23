<!DOCTYPE html>
	<!--#include virtual="includes/connection.asp"-->
    <!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->

<%
	requestClass = RequestParse(Request("Class"),"")	
	requestServer = RequestParse(Request("Server"),"")	

    Title = "Database Classes - Table Report"

	Dim rs: Set rs = Server.CreateObject("ADODB.Recordset")
		
	If requestClass > "" Then
		SQL="ha_DatabaseClasses_read('" & requestClass & "')"
		rs.open SQL, connLocal, 3, 3, 4

		If rs.eof Then
			requestClass = ""
			dbName = ""
		Else
			dbName = rs("DatabaseName")
		End If
		rs.close
	End If
%>

<html>
<head>
    <title><%=Title%></title>

    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

    <link href="/ha_Backoffice/includes/ha_BackOfficePopup.css" type="text/css" rel="stylesheet">
    <script language="javascript" type="text/javascript" src="/includes/haFunctions.js"></script>
	
	<script type="text/javascript">
	
	function pageReload(ID)
        {
		var e = document.getElementById('ClassName');
		var selectedClass = e.options[e.selectedIndex].text;
		var e = document.getElementById('ServerName');
		var selectedServer = e.options[e.selectedIndex].text;
		
        var URL = 'ha_Table_Report.asp' +
            '?Class=' + selectedClass +
			'&Server=' + selectedServer;
        window.location = URL;
        }
		
		
	</script>

	<style type="text/css">
		td
			{
			font-size: 10pt;
			}
			
		td.head
			{
			font-weight: bold;
			font-size: 8pt;
			font-stretch: condensed;
			border-bottom: 1px solid black;
			}

	</style
	
</head>


<body onLoad="SetScreen(1000,850);">

  <span class="popupHeading"><%=Title%></span>

<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td>
	  
	  <table width="95%" border="0" cellspacing="2" cellpadding="2" align="center">
		<tr><td>&nbsp;</td></tr>
		
		<tr>
		  <td width="50%">
			<b>Class Name:</b>
			<Select onchange="pageReload();" id="ClassName" name="ClassName">
  
<%
					strSQL = "util_HTML_option_list('ha_Database_Classes', '', '', 'ClassName', 'ClassName', '', '', '" & requestClass & "')"
					rs.Open strSQL, connLocal, 3, 3, 4
					If not rs.eof Then
						While not rs.eof
							Response.write rs("result")
							rs.MoveNext
						Wend
					End If
					rs.Close
%>
		
			</Select>
			<%=dbName%>
		  </td>
			  
		  <td width="50%">
			<b>Server Name:</b>
			<Select onchange="pageReload();" id="ServerName" name="ServerName">
  
<%
					strSQL = "util_HTML_option_list('ha_Server_Info', 'isnumeric(DropDownSequence)', '1', 'ServerName', 'ServerName', 'DropDownSequence', '', '" & requestServer & "')"
					rs.Open strSQL, connLocal, 3, 3, 4
					If not rs.eof Then
						While not rs.eof
							Response.write rs("result")
							rs.MoveNext
						Wend
					End If
					rs.Close
%>
		
			</Select>
		  </td>
			  
		</tr>
		<tr><td>&nbsp;</td></tr>
	  </table>
	</td>
  </tr>
	
  <tr>
	<td>
	  <table width="95%" border="0" cellspacing="2" cellpadding="2" align="center">
	
		<tr>
		  <td width="40%" class="head">&ensp;Table Name</td>
		  <td width="15%" class="head" align="center">In Database</td>
		  <td width="15%" class="head" align="center">In Schema</td>
		  <td width="15%" class="head" align="center">Row Count</td>
		  <td width="15%" class="head" align="center">Disk Space (KBs)</td>
		<tr>
		
		<tr>
		  <td colspan="5">
			<div style="width: 928px; height: 530px; overflow-y: scroll;">
			  <table width="100%" border="0" cellspacing="2" cellpadding="2">
<%
					strSQL = "ha_Database_Report('" & requestClass & "', '" & requestServer & "')"
					rs.Open strSQL, connLocal, 3, 3, 4
					If not rs.eof Then
						recordCount = 0
						While not rs.eof

							'Dividing Separator
							recordCount = recordCount + 1
							If recordCount mod 3 = 0 Then
								detailStyle = "border-bottom: 1px solid #C0C0C0;"
							Else
								detailStyle = "border: 0;"
							End If
							
							'Add color for bad records
							If rs("DisplayInDatabase") = "No" or rs("DisplayInSchema") = "No" Then
								detailStyle = detailStyle & " color: red;"
							Else
								detailStyle = detailStyle & " color: black;"
							End If
							
							If Left(rs("TableName"),11) = "No Database" Then
%>
				<tr>
				  <td width="100%" align="center" style="color: red;">
				    <br><br><br><br><br><br><br>
					<%=rs("TableName")%>
				  </td>
				</tr>
<%
							Else
%>
				<tr>
				  <td width="40%" style="<%=detailStyle%>">
					<%=rs("TableName")%>
				  </td>
				  <td width="15%" style="<%=detailStyle%>" align="center">
					&ensp;<%=rs("DisplayInDatabase")%>
				  </td>
				  <td width="15%" style="<%=detailStyle%>" align="center">
					&ensp;<%=rs("DisplayInSchema")%>
				  </td>
				  <td width="15%" style="<%=detailStyle%>" align="right">
					<%=rs("DisplayRowCount")%>&emsp;&emsp;&emsp;
				  </td>
				  <td width="15%" style="<%=detailStyle%>" align="right">
					<%=rs("DisplayTotalSpace")%>&emsp;&emsp;
				  </td>
				</tr>
					
<%					
							End If
							rs.MoveNext
						Wend
					End If
					rs.Close
%>
	
	
			  </table>
			</div>
		  </td>
		</tr>
	  
	  </table>
	</td>
  </tr>
  
</table>

  <span class="popupTail">
	&nbsp;
  </span>
  
  <span class="popupButtons">
	<a href="javascript:window.close();" class="bo_Button80">Close</a>
  </span>


</body>
</html>
