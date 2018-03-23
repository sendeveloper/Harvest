<%
	Session("Redirect") = ""
	ColorTab = 1
	PageHeading = "Server Archives"
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

    <style type="text/css">

		th.Heading1
			{
			font-size:			14px;
			text-align:			center;
			border-bottom:		1px solid black;
			}
    </style>

    <script language="javascript" type = "text/javascript">
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

			  <tr valign="top">
				<th class="Heading1">&nbsp;</th>
				<th class="Heading1" colspan="2">Source</th>
				<th class="Heading1" colspan="2">Archive 1</th>
				<th class="Heading1" colspan="2">Archive 2</th>
			  </tr>

			  <tr style="background-color: #008000">
				<th style="color: White">Client</th>
				<th style="color: White">Location</th>
				<th style="color: White">Object</th>
				<th style="color: White">Location</th>
				<th style="color: White">Last Date</th>
				<th style="color: White">Location</th>
				<th style="color: White">Last Date</th>
			  </tr>
			
	<%
		recordCount = 0
		LineCount = 0

		set rs=server.createObject("ADODB.Recordset")
		SQL = "SELECT * FROM ha_Archive_info WHERE Class = 'Archive' ORDER BY SourceClient"
		rs.Open SQL, connCasper10, 2, 3 'connDallas

		if not RS.EOF then
			do while not rs.eof
				recordCount = recordCount + 1
				LineCount = LineCount + 1
				If LineCount > 2 then
				  LineCount = 0
				  strColor = "#FFFFCC"
				Else
				  strColor = "#FFFFFF"
				End If
	%>
			
			  <tr bgcolor=<%=strColor%>>
				<td>
				  <%=rs("SourceClient")%>
				</td>
				<td>
				  <%=rs("SourceLocation")%>
				</td>
				<td>
				  <%=rs("SourceObject")%>
				</td>
				<td>
				  <%=rs("Dest1Location")%>
				</td>
				<td>
				  <%=rs("Dest1LastAction")%>
				</td>
				<td>
				  <%=rs("Dest2Location")%>
				</td>
				<td>
				  <%=rs("Dest2LastAction")%>
				</td>
			  </tr>
			  
	<%
				rs.MoveNext
			loop
		end if
		
		rs.close
	%>

			  <tr>
				<th style="background-color: #008000" colspan="7">&nbsp;</th>
			  </tr>
			  
			  <tr><td>&nbsp;</td></tr>
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

