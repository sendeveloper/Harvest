<%
	Session("Redirect") = ""
	ColorTab = 1
	PageHeading = "Table Distribution"
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
			text-align:			center;
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
				<th class="Heading1" colspan="2">Source</th>
				<th class="Heading1">&nbsp;</th>
				<th class="Heading1" colspan="3">Last Run</th>
				<th class="Heading1">&nbsp;</th>
			  </tr>

			  <tr>
				<th width="10%">Table</th>
				<th width="20%">Location</th>
				<th width="10%">Frequency</th>
				<th width="10%">Date</th>
				<th width="10%">Time</th>
				<th width="20%">Duration Since Last</th>
				<th width="20%">Action</th>
			  </tr>
			
	<%
		recordCount = 0
		LineCount = 0

		set rs=server.createObject("ADODB.Recordset")
		SQL = "ha_TableDistribution_read"
		rs.Open SQL, objcon, 3, 3, 4 'connDallas

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
			
			  <tr>
				<td>
				  <%=rs("TableName")%>
				</td>
				<td>
				  <%=rs("TableLocation")%>
				</td>
				<td>
				  <%=rs("Frequency")%>
				</td>
				<td>
				  <%=rs("LastRunDate")%>
				</td>
				<td>
				  <%=rs("LastRunTime")%>
				</td>
				<td>
				  <%=rs("LastRunDuration")%>
				</td>
				<td>
				  <%=rs("ActionButton")%>
				</td>
			  </tr>
			  
	<%
				rs.MoveNext
			loop
		end if
		
		rs.close
		objcon.close 'connDallas
	%>

			  <tr>
				<th class="Heading1" colspan="7">&nbsp;</th>
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

