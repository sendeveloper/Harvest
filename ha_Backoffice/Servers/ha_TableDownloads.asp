<!DOCTYPE html>
<!--#include virtual="includes/connection.asp"-->
<!--#include virtual="ha_BackOffice/includes/Config.asp"-->
<!--#include virtual="ha_backoffice/includes/FunctionsTimeDifference.inc"-->

<%
	ColorTab = 1
	PageHeading = "Table Downloads"
%>

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<title>Harvest American Backoffice - <%=PageHeading%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    
	<link href="<%=strPathIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet" />
	<link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>menu/dropdowntabfiles/halfmoontabs.css" />
	<script type="text/javascript" src="<%=strPathIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>
  	<script type="text/javascript" src="<%=strPathBOIncludes%>haFunctions.js" language="javascript"></script>
	
    <style type="text/css">
		td
			{
			vertical-align:	top;
			}		
			
		th
			{
			border-bottom:	1px solid black;
			}
			
    </style>
</head>

<body class="gray_desktop">
  <table class="MainBodyFixed">	
    <tr>
      <td>
        <!--#include virtual="ha_Backoffice/includes/bodyparts/ha_PageHeading.inc"-->
        <table class="MainTable">
	
		  <tr><td>&nbsp;</td></tr>
	
		  <tr>
			<td align="center">
			  <table width="95%" cellpadding="2" cellspacing="2">
				<tr>
				  <td>
					For Zip2Tax, csv tables are downloaded from a choice of locations.  The history is written immediately to the local
					copy of the table download log.  Eventually that history is transferred to the Master table on Philly01 and then
					back to replication tables on each local server.
				  </td>
				</tr>

				<tr><td>&nbsp;</td></tr>
		  
				<tr>
				  <td align="right">
					<a href="javascript:clickPopupImage(1554);" class="TopRightLink">Customer Request Flowchart</a>
					<a href="javascript:clickPopupImage(1555);" class="TopRightLink">Background Process Flowchart</a>
					<a href="javascript:clickPopupImage(1483);" class="TopRightLink">Report Process Flowchart (Yet to be done)</a>
				  </td>
				</tr>
		  
				<tr><td>&nbsp;</td></tr>
		  
				  <tr>
					<td align="Center">
					  <table width="90%" cellpadding="2" cellspacing="2">
				  
				  
					  <tr>
						
						<th colspan="7">Table Download Count</th>
					  </tr>
					  
					  <tr>
						<th width="20%" align="left">Table Type</th>
						<th width="20%" align="left">Server Name</th>
						<th width="12%">Today</th>
						<th width="12%">Yesterday</th>
						<th width="12%">This Month</th>
						<th width="12%">This Year</th>
						<th width="12%">Total</th>
					  </tr>
		  
        
<%
		
		set rs=server.createObject("ADODB.Recordset")
		SQL = "SELECT * FROM z2t_BackOffice.dbo.z2t_TableDownload_report ORDER BY Sequence"
		rs.Open SQL, connLocal
		
		If not RS.EOF Then
			
			do while not rs.eof
				'Skip a row
				If previousTableType <> rs("NameType") Then
					Response.Write "<tr><td>&nbsp;</td></tr>"
				End If
				
%>
			  <tr>
<%
				If previousTableType <> rs("NameType") Then
					Response.Write "<td>" & rs("NameType") & "</td>"
				Else
					Response.Write "<td>&nbsp;</td>"
				End If
%>
				<td><%=rs("NameServer")%></td>
				<td align="right"><%=DisplayCount(rs("countToday"))%></td>
				<td align="right"><%=DisplayCount(rs("countYesterday"))%></td>
				<td align="right"><%=DisplayCount(rs("CountThisMonth"))%></td>
				<td align="right"><%=DisplayCount(rs("CountThisYear"))%></td>
				<td align="right"><%=DisplayCount(rs("CountTotal"))%></td>
			  </tr>
<%
				previousTableType = rs("NameType")				
				AsOf = rs("EditedDate")
				rs.MoveNext
			loop	
		End If
		
		rs.close

		
Function DisplayCount(count)
	DisplayCount = FormatNumber(count,0) & "&emsp;&emsp;"
End Function

%>
					  <tr>						
						<th colspan="7">&nbsp;</th>
					  </tr>
		  
					  <tr>
						<td colspan="3">
						  Last Updated: <%=AsOf%> -- <%=TimeDifference(AsOf, Now)%>
						</td>
					  </tr>
				
					</table>
				  </td>
				</tr>
										  
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

