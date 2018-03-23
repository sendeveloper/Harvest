<!DOCTYPE html>
<!--#include virtual="includes/connection.asp"-->
<!--#include virtual="ha_BackOffice/includes/Config.asp"-->

<%
	ColorTab = 10
	PageHeading = "Order Upload Process"
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
					On-line orders are sent to us from the Shopping Cart.  We collect this data and process it.
				  </td>
				</tr>

				<tr><td>&nbsp;</td></tr>
		  
				<tr>
				  <td>
					1.) Pass HTTP datastream
					  <a href="javascript:clickPopupImage(1484);" class="TextInlineLink">WooCommerce Extra Code</a>
					  <a href="javascript:clickPopupImage(1485);" class="TextInlineLink">WooCommerce XML Export Suite Settings</a>
					  <a href="javascript:clickPopupImage(1488);" class="TextInlineLink">WooCommerce Order Upload Flowchart</a><br>
					2.) Parse order XML data received
					  <a href="javascript:clickPopupImage(1489);" class="TextInlineLink">WooCommerce XML Data Parse</a><br>
					3.) Place into Account, Order and Order Details tables
					  <a href="javascript:clickPopupImage(1490);" class="TextInlineLink">WooCommerce Order Process</a><br>
					4.) Handle Requirements
					  <a href="javascript:clickPopupImage(1491);" class="TextInlineLink">Order Requirements Checking</a>
					  <a href="javascript:clickPopupImage(1492);" class="TextInlineLink">Order Requirements Fulfillment</a><br>
				  </td>
				</tr>
				
				<tr><td>&nbsp;</td></tr>
				
				<tr>
				  <td>
					<hr>
				  </td>
				</tr>

				<tr><td>&nbsp;</td></tr>
				
				<!--Requirements-->
				<tr>
				  <td class="GreenHead">Requirements</td>
				</tr>
				
				<tr>
				  <td style="padding-left: 2em;">
					<table width="100%" cellpadding="2" cellspacing="2">
					
					  <tr>
						<th>
						  Requirement Name
						</th>
						<th>
						  Purpose
						</th>
						<th>
						  Products Affected
						</th>
						<th>
						  Actions
						</th>
					  </tr>
					  <tr>
						<td>
						  Zip2Tax Lookup Process
						</td>
					  </tr>
					  <tr>
						<td>
						  Zip2Tax Lookup E-mail
						</td>
					  </tr>
					  <tr>
						<td>
						  Zip2Tax Link Process
						</td>
					  </tr>
					  <tr>
						<td>
						  Zip2Tax Initial Table Process
						</td>
					  </tr>
					  <tr>
						<td>
						  Zip2Tax Shipping
						</td>
					  </tr>
					</table>
				  </td>
				</tr>
		  
				<tr>
				  <td>
					<hr>
				  </td>
				</tr>
			  
				<tr><td>&nbsp;</td></tr>
		  
				<!--Order Counts-->
				<tr>
				  <td class="GreenHead">Order Counts</td>
				</tr>
				
				<tr>
				  <td style="padding-left: 2em;">
					<table width="100%" cellpadding="2" cellspacing="2">
				    
					  <tr>
						<td>&nbsp;</td>
						<th colspan="5">Order Counts</th>
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
		Dim connCasper08: Set connCasper08 = Server.CreateObject("ADODB.Connection")
		connCasper08.Open "driver=SQL Server;server=10.119.55.116,7843;uid=davewj2o;pwd=get2it;database=ha_Accounting"  'Casper08
		
		set rs=server.createObject("ADODB.Recordset")
		SQL = "ha_Order_counts"
		rs.Open SQL, connCasper08, 3, 3, 4

		If not rs.eof Then
			Do while not rs.eof
				'Skip rows
				If rs("OperationName") = "WooCommerce Upload Raw" _
					or rs("OperationName") = "Orders - Manual" Then
					Response.Write "<tr><td>&nbsp;</td></tr>"
				End If
				
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
		connCasper08.close
		
Function DisplayCount(count)
	DisplayCount = FormatNumber(count,0) & "&emsp;&emsp;"
End Function

%>
		  
					</table>
				  </td>
				</tr>
								  
				<tr>
				  <td colspan="4">
					<hr>
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

