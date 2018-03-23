<%
	Session("Redirect") = ""
	ColorTab = 1
	PageHeading = "Table Distribution - Q Process"
%>

<!--#include virtual="includes/connection.asp"-->
<!--#include virtual="ha_backoffice/includes/Config.asp"-->

<html>
<head>
    <title>Harvest American Backoffice - <%=PageHeading%></title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
  
    <link href="<%=strPathIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet">
    <link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>menu/dropdowntabfiles/halfmoontabs.css">

    <script type="text/javascript" src="<%=strPathIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>
    <script type="text/javascript" src="<%=strPathBOIncludes%>haFunctions.js" language="javascript"></script>

    <style type="text/css">
		th
			{
			text-align:			left;
			border-bottom:		1px solid black;
			}

		th.Heading1
			{
			font-size:			14px;
			text-align:			left;
			border-bottom:		1px solid black;
			}

    </style>
	
	<script language="javascript" type = "text/javascript">

        function clickPopup(n)
            {
            var URL = n + '.asp'
            openPopUp(URL);
            }
			
    </script>

</head>

<body class="gray_desktop">
  <table align="center" cellspacing="0" cellpadding="0" class="MainBody">
    <tr>
	  <td>

		<!--#include virtual="ha_Backoffice/includes/bodyparts/ha_PageHeading.inc"-->

				<table width="100%" border="0" cellspacing="0" cellpadding="0">
	
		<tr><td>&nbsp;</td></tr>
		
		<tr>
		  <td align="center">
			<table width="95%" border="0" cellspacing="5" cellpadding="5">

			  <tr vAlign="top">
				<th class="Heading1" colspan="8">Fill The Q</th>
			  </tr>

			  <tr>
				<th width="10%">Server</th>
				<th width="10%">Database</th>
				<th width="10%">Table Name</th>
				<th width="10%">Q Name</th>
				<th width="15%" style="text-align: center;">Entries<br>Added<br>Today</th>
				<th width="15%" style="text-align: center;">Entries<br>Processed<br>Today</th>
				<th width="15%" style="text-align: center;">Entries<br>Added<br>Last 30 Days</th>
				<th width="15%" style="text-align: center;">Entries<br>Processed<br>Last 30 Days</th>
			  </tr>

	<%
		recordCount = 0

		set rs=server.createObject("ADODB.Recordset")
		SQL = "ha_Table_Distribution_Q_Process_list"
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
				  <%=rs("NameServer")%>
				</td>
				<td style="<%=detailStyle%>">
				  <%=rs("NameDatabase")%>
				</td>
				<td style="<%=detailStyle%>">
				  <%=rs("NameTable")%>
				</td>
				<td style="<%=detailStyle%>">
				  <%=rs("NameQ")%>
				</td>
				<td style="<%=detailStyle%> text-align: center;">
				  <%=rs("TodayCountEntries")%>
				</td>
				<td style="<%=detailStyle%> text-align: center;">
				  <%=rs("TodayCountProcessed")%>
				</td>
				<td style="<%=detailStyle%> text-align: center;">
				  <%=rs("TotalCountEntries")%>
				</td>
				<td style="<%=detailStyle%> text-align: center;">
				  <%=rs("TotalCountProcessed")%>
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
			  


			  <tr vAlign="top">
				<th class="Heading1" colspan="8">Process The Q</th>
			  </tr>
			  
			  <tr>
				<th colspan="4">Triggered By</th>
				<th Colspan="4">Q Location/Name</th>
			  </tr>
			  
  			  <tr>
			    <td colspan="4">Casper10.ha_Backoffice.ha_Processes_min</td>
				<td colspan="4">Casper10.ha_Backoffice.ha_Q</td>
			  </tr>

			  <tr><td>&nbsp;</td></tr>


			  <tr vAlign="top">
				<th class="Heading1" colspan="8">What does Q Process do?</th>
			  </tr>
			  
  			  <tr>
				<td colspan="8">
					Reads in the entries in the particular Q you're working on<br>
					Marks those entries as being worked on<br>
					Calls the appropriate table distribution asp for each Q Class<br>
					Marks the entries that were completed as Deleted<br>
					Releases any marks it may have had on Q entires it didn't complete<br>
					<br>
					Each Q is purged daily removing anything over 30 days old by ha_Processes_day<br>
				</td>
			  </tr>
			  
			  <tr><td>&nbsp;</td></tr>
			  
			  
			  
			  <tr>
				<th class="Heading1" colspan="8">&nbsp;</th>
			  </tr>
			  
			  <tr><td>&nbsp;</td></tr>
			  
			</table>
		  </td>
		</tr>
		</table>

	  </td>
	</tr>
  </table>
  <!--#include virtual="ha_BackOffice/includes/BodyParts/ha_PageTail.inc"-->
</body>
</html>


