<!DOCTYPE html>
<!--#include virtual="includes/connection.asp"-->
<!--#include virtual="ha_BackOffice/includes/Config.asp"-->

<%
	ColorTab = 1
	PageHeading = "Load Balancer"
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
			
		div.members
			{
			border:			1px solid gray;
			height:			7em;
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
					To serve the Zip2Tax.com api we have dedicated Philly02 and Philly04 to this purpose.  They are connected in a Load Balanced configuration.
					Casper06 also serves the api to the West Coast.
				  </td>
				</tr>

				<tr><td>&nbsp;</td></tr>
		  
				<tr>
				  <td align="right">
					<a href="javascript:clickPopupImage(1443);" class="TopRightLink">Load Balancer Overview</a>
				  </td>
				</tr>
		  
				<tr><td>&nbsp;</td></tr>
		  
				  <tr>
					<td align="Center">
					  <table width="90%" cellpadding="2" cellspacing="2">
				  
				  
						<tr>
						  <td colspan="4">&nbsp;</td>
						  <th colspan="4">Lookup Count</th>
						</tr>
						  
						<tr>
						  <th width="14%">Server Name</th>
						  <th width="14%">Server IP</th>
						  <th width="14%">Load Balancer IP</th>
						  <td width="2%">&nbsp;</td>				
						  <th width="14%">Today</th>
						  <th width="14%">Yesterday</th>
						  <th width="14%">This Month</th>
						  <th width="14%">This Year</th>
						</tr>
		  
<%		  
		Dim connz2tBackoffice: Set connz2tBackoffice = Server.CreateObject("ADODB.Connection")
		'connz2tBackoffice.Open "driver=SQL Server;server=66.119.50.230,7043;uid=davewj2o;pwd=get2it;database=z2t_Backoffice"  'Casper10
		connz2tBackoffice.Open "driver=SQL Server;server=127.0.0.1,7043;uid=davewj2o;pwd=get2it;database=z2t_Backoffice"  'Casper10
		
		set rs=server.createObject("ADODB.Recordset")
		SQL = "z2t_Activity_counts_api"
		rs.Open SQL, connz2tBackoffice, 3, 3, 4

		If not RS.EOF Then
			Do while not rs.eof
%>
			  <tr>
				<td><%=rs("ServerName")%></td>		
				<td align="center"><%=rs("ServerIP")%></td>		
				<td align="center"><%=rs("LoadBalancerIP")%></td>					
				<td>&nbsp;</td>
				<td align="right"><%=FormatNumber(rs("countDay"),0)%>&emsp;&emsp;</td>		
				<td align="right"><%=FormatNumber(rs("countYesterday"),0)%>&emsp;&emsp;</td>		
				<td align="right"><%=FormatNumber(rs("countMonth"),0)%>&emsp;&emsp;</td>		
				<td align="right"><%=FormatNumber(rs("countYear"),0)%>&emsp;&emsp;</td>		
			  </tr>
<%
				rs.MoveNext
			Loop
		End If
		
		rs.close
		connz2tBackoffice.close
%>
					<tr><td>&nbsp;</td></tr>
					
					<tr>
					  <td colspan="4">&nbsp;</td>
					  <td colspan="4" align="center"><i>(Data can be delayed by as much as 15 minutes)</i></td>
					</tr>
		  
					</table>
				  </td>
				</tr>
				
				<tr><td>&nbsp;</td></tr>
								  
				<tr>
				  <td colspan="4">
					<hr>
				  </td>
				</tr>
		  
		  
				<tr><td>&nbsp;</td></tr>
		  
				<tr>
				  <td>
					<b>API Formats</b>
				  </td>
				</tr>
		  
				<tr><td>&nbsp;</td></tr>
		  
				  <tr>
					<td align="Center">
					  <table width="90%" cellpadding="2" cellspacing="2">
				  						  
						<tr>
						  <th width="40%">API Name</th>
						  <th width="20%">Subscription Required</th>
						  <th width="40%">Used By</th>
						</tr>
		  
						<tr>
						  <td>api.Zip2Tax.com/TaxRate-USA.json</td>
						  <td>Database Interface</td>
						  <td>Customers</td>
						</tr>
						
						<tr>
						  <td>api.Zip2Tax.com/TaxRate-USA.xml</td>
						  <td>Database Interface</td>
						  <td>Customers</td>
						</tr>
						
						<tr>
						  <td>api.Zip2Tax.com/TaxRate-USA.raw</td>
						  <td>Database Interface</td>
						  <td>Don't know?</td>
						</tr>
						
						<tr>
						  <td>api.Zip2Tax.com/TaxRate-USA.web</td>
						  <td>Online Lookup</td>
						  <td>Website?</td>
						</tr>
						
						<tr>
						  <td>api.Zip2Tax.com/TaxRate-USA.widget</td>
						  <td>Online Lookup</td>
						  <td>Widget, iPhone?, Android?</td>
						</tr>
						
						<tr>
						  <td>api.Zip2Tax.com/TaxRate-USA.widgetlogin</td>
						  <td>Online Lookup</td>
						  <td>Don't know?</td>
						</tr>
						
						<tr>
						  <td>api.Zip2Tax.com/TaxRate-USA.webxml</td>
						  <td>Online Lookup</td>						  
						  <td>Don't know?</td>
						</tr>
		  
					</table>
				  </td>
				</tr>
				
				<tr><td>&nbsp;</td></tr>
				  
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

