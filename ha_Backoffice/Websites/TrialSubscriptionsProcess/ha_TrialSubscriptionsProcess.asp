<!DOCTYPE html>
<!--#include virtual="includes/connection.asp"-->
<!--#include virtual="ha_BackOffice/includes/Config.asp"-->

<%
	ColorTab = 10
	PageHeading = "Trial Subscriptions Process"
	
	Indent = "&emsp;&emsp;&emsp;&emsp;"
%>

<html>
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
			
		td.GreenHead
			{
			font-size:		11pt;
			font-weight:	bold;
			color:			green;
			}
			
		th
			{
			border-bottom:	1px solid black;
			}			
    </style>
</head>

<body class="gray_desktop">
  <table class="MainBodyFixed">	
    <tr valign="top">
      <td>
        <!--#include virtual="ha_Backoffice/includes/bodyparts/ha_PageHeading.inc"-->
        <table class="MainTable">
	
		  <tr><td>&nbsp;</td></tr>
	
		  <tr>
			<td align="center">
			  <table width="95%" cellpadding="2" cellspacing="2">
				<tr>
				  <td>
					Prospects are encouraged to sign up for free trials of our lookup service.  
					We perform this in two steps.
				  </td>
				</tr>

				<tr><td>&nbsp;</td></tr>
		  
				<tr>
				  <td>
					1.) Step 1: Collect basic information, create an interim account and send an E-mail.<br>
					  <%=Indent%><a href="javascript:clickPopupImage(1511);" class="TextInlineLink">Website Storyboard</a><br>
					  <%=Indent%><a href="javascript:clickPopupImage(1513);" class="TextInlineLink">Data Process</a><br><br>
					  
					2.) Step 2: After sign-in with an interim account additional data is collected.<br>
					  <%=Indent%><a href="javascript:clickPopupImage(1512);" class="TextInlineLink">Website Storyboard</a><br>
					  <%=Indent%><a href="javascript:clickPopupImage(1514);" class="TextInlineLink">Data Process</a><br><br>
				  </td>
				</tr>
				
				<tr><td>&nbsp;</td></tr>
				
				<tr>
				  <td>
					<hr>
				  </td>
				</tr>
				
			  </table>
			</td>
		  </tr>
		  
		  <tr>
			<td align="center">
			  <table width="95%" cellpadding="2" cellspacing="2">
				
				<!--Order Counts-->
				<tr>
				  <td class="GreenHead">Trial Subscriptions Counts</td>
				</tr>
				
				<tr>
				  <td style="padding-left: 2em;">
					<table width="100%" cellpadding="2" cellspacing="2">
				    
					  <tr>
						<td>&nbsp;</td>
						<th colspan="5">Trial Subscriptions Counts</th>
					  </tr>
					  
					  <tr>
						<th width="20%">Source</th>
						<th width="16%">Today</th>
						<th width="16%">Yesterday</th>
						<th width="16%">This Month</th>
						<th width="16%">This Year</th>
						<th width="16%">Total</th>
					  </tr>
		  
<%		  
		Dim connPhilly01: Set connPhilly01 = Server.CreateObject("ADODB.Connection")
		connPhilly01.Open "driver=SQL Server;server=208.88.49.18,8143;uid=davewj2o;pwd=get2it;database=z2t_Subscriptions"  'Philly01
		
		set rs=server.createObject("ADODB.Recordset")
		SQL = "z2t_Subscriptions_Interim_counts"
		rs.Open SQL, connPhilly01, 3, 3, 4

		If not rs.eof Then
			Do while not rs.eof				
%>

					  <tr>
						<td><%=rs("OperationName")%></td>
						<td align="right"><%=DisplayCount(rs("countDay"))%></td>
						<td align="right"><%=DisplayCount(rs("countYesterday"))%></td>
						<td align="right"><%=DisplayCount(rs("countMonth"))%></td>
						<td align="right"><%=DisplayCount(rs("countYear"))%></td>
						<td align="right"><%=DisplayCount(rs("countTotal"))%></td>
					  </tr>
			  
<%
				rs.MoveNext
			Loop
		End If
		
		rs.close
		connPhilly01.close
		

Function DisplayCount(count)
	DisplayCount = FormatNumber(count,0) & "&emsp;&emsp;"
End Function

%>

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

