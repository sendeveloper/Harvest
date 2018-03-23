<!DOCTYPE html>
<!--#include virtual="includes/connection.asp"-->
<!--#include virtual="ha_BackOffice/includes/Config.asp"-->

<%
	ColorTab = 10
	PageHeading = "WordPress"
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
					WordPress is a CMS (Content Management System) used for web development.
					A semantic personal publishing platform with a focus on aesthetics, web standards, and usability.
				  </td>
				</tr>

				<tr><td>&nbsp;</td></tr>
		  
				<tr>
				  <td align="right">
					<a href="javascript:clickPopupImage(1587);" class="TopRightLink">WordPress Overview</a>
				  </td>
				</tr>
		  
				<tr><td>&nbsp;</td></tr>
		  
		  <tr>
		    <td align="Center">
			  <table width="80%" cellpadding="2" cellspacing="2">
		  
		  
				<tr>
				  <th width="30%">Domain Name</th>
				  <th width="30%">Subdomain Name</th>
				  <th width="10%">Host</th>
				  <th width="10%">SSL</th>				
				  <th width="20%">Purpose</th>
				</tr>
				
				<tr>
				  <td>blog.RaffleTicketSoftware.com</td>
				  <td>&nbsp;</td>
				  <td align="center">QTH</td>
				  <td align="center">No</td>
				  <td>RTS blog</td>
				</tr>
		  
				<tr>
				  <td>&nbsp;</td>
				  <td>blog.Zip2Tax.com</td>
				  <td align="center">QTH</td>
				  <td align="center">No</td>
				  <td>Z2T blog</td>
				</tr>
				
				<tr>
				  <td colspan="5">
				    <hr>
				  </td>
				</tr>
				
				<tr>
				  <td>www.HarvestAmerican.com</td>
				  <td>&nbsp;</td>
				  <td align="center">QTH</td>
				  <td align="center">No</td>
				  <td>E-mail only</td>
				</tr>
				
				<tr>
				  <td colspan="5">
				    <hr>
				  </td>
				</tr>
				
				<tr>
				  <td>www.NumberMachinePro.com</td>
				  <td>&nbsp;</td>
				  <td align="center">QTH</td>
				  <td align="center">No</td>
				  <td>NMP Website</td>
				</tr>
				
				<tr>
				  <td colspan="5">
				    <hr>
				  </td>
				</tr>
				
				<tr>
				  <td>wp.Zip2Tax.com</td>
				  <td>&nbsp;</td>
				  <td align="center">QTH</td>
				  <td align="center">No</td>
				  <td>Z2T Website (future)</td>
				</tr>
								
				<tr>
				  <td colspan="5">
				    <hr>
				  </td>
				</tr>
				
				<tr>
				  <td>beta.RaffleTicketSoftware.com</td>
				  <td>&nbsp;</td>
				  <td align="center">WP Engine</td>
				  <td align="center">Yes</td>
				  <td>RTS Website (future)</td>
				</tr>
				
				<tr>
				  <td colspan="5">
				    <hr>
				  </td>
				</tr>
				
				<tr>
				  <td>beta.RaffleTicketSoftware.com</td>
				  <td>&nbsp;</td>
				  <td align="center">Go Daddy</td>
				  <td align="center">Yes</td>
				  <td>RTS Website</td>
				</tr>
				
				<tr>
				  <td colspan="5">
				    <hr>
				  </td>
				</tr>
				
				<tr>
				  <td>www.RaffleTicketSoftware.com</td>
				  <td>&nbsp;</td>
				  <td align="center">Casper 7</td>
				  <td align="center">&nbsp;</td>
				  <td>RTS Website</td>
				</tr>
				<tr>
				  <td colspan="5">
				    <hr>
				  </td>
				</tr>
				
				<tr>
				  <td>www.OurRafflePage.com</td>
				  <td>&nbsp;</td>
				  <td align="center">WP Engine</td>
				  <td align="center">&nbsp;</td>
				  <td>RTS Website</td>
				</tr>
				<tr>
				
				  <td colspan="5">
				    <hr>
				  </td>
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

