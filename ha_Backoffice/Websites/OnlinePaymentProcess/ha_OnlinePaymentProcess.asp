<!DOCTYPE html>
<!--#include virtual="includes/connection.asp"-->
<!--#include virtual="ha_BackOffice/includes/Config.asp"-->

<%
	ColorTab = 10
	PageHeading = "Online Payment Process"
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
					We have the ability to create invoices and direct customers to an on-line location where they
					can pay for them using a Cybersource application.<br><br>     
					
					This process is different and shoudn't be confused with our shopping cart processes.
				  </td>
				</tr>

				<tr><td>&nbsp;</td></tr>
		  
				<tr>
				  <td align="right">
					<a href="javascript:clickPopupImage(1542);" class="TopRightLink">Payment Process Overview</a>
				  </td>
				</tr>
		  
				<tr><td>&nbsp;</td></tr>
		  
		
				<tr>
                	<td>
                    	<b>1)- Add New Order</b>
                        	(To create new invoice for customer, pay back or pay tax) from number-it.com Backoffice)
                    </td>
                </tr>
                
                <tr>
                	<td>
                    	&nbsp;&nbsp;<b><a href="javascript:clickPopupImage(1545);" class="TopRightLink">Process Steps</a></b> &nbsp
                		<b><a href="javascript:clickPopupImage(1545);" class="TopRightLink">Process Flow</a></b> &nbsp
                        <b><a href="javascript:clickPopupImage(1545);" class="TopRightLink">Data Flow</a></b> &nbsp
                    
                    </td>
                </tr>
			  
              	<tr>
                	<td>
                    	<b>2)- Online Payment Process (Customer)</b>
                        	(To pay invoice from (www,dev,staging).Zip2Tax.com)
                    </td>
                </tr>
                
                <tr>
                	<td>
                    	&nbsp;&nbsp;<b><a href="javascript:clickPopupImage();" class="TopRightLink">Process Steps</a></b> &nbsp
                		&nbsp;&nbsp;<b><a href="javascript:clickPopupImage();" class="TopRightLink">Process Flow</a></b> &nbsp
                        &nbsp;&nbsp;<b><a href="javascript:clickPopupImage();" class="TopRightLink">Data Flow</a></b> &nbsp
                    
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

