<!--#include virtual="ha_backoffice/includes/Config.asp"-->
<!--#include virtual="includes/connection.asp"-->

<%
	ColorTab = 1
	PageHeading = "Shopping Carts Process" 
%>

<html>
<head>
    <title>Harvest American Backoffice - Shopping Carts Process</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />

    <link href="<%=strPathBOIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet" />
    <link rel="stylesheet" type="text/css" href="<%=strPathBOIncludes%>menu/dropdowntabfiles/halfmoontabs.css" />

    <script type="text/javascript" src="<%=strPathBOIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>
	<script type="text/javascript" src="<%=strPathBOIncludes%>haFunctions.js" language="javascript"></script>
	<script language='JavaScript'>

	function clickPopup(ID)
		{
			var URL = '/ha_Backoffice/Company/DocumentMaintenance/ha_Document_Show.asp' +
			'?PhotoID=' + ID;
			window.open(URL,'','height=490, width=600')
		}
		
	</script>
    <style type="text/css">
		span.subHead
			{
			font-weight: 	bold;
			color:			#336600;
			}
		span.newStyleHeading
			{
			font-size: 		11pt;
			font-weight: 	bold;
			color:			#336600;
			}	
    </style>

</head>

<body class="gray_desktop">
  <table align="center" cellspacing="0" cellpadding="0" class="MainBody">
    <tr>
		<td>

			<!--#include virtual="ha_Backoffice/includes/bodyparts/ha_PageHeading.inc"-->
			<center>
			<table width="96%" cellspacing="3" cellpadding="5" class="section">
		<tr>
			<td align="center">
			<center><b><u><a href="https://www.snscart.com/admin/">Raffle Ticket</a></u></b></td>
			<td align="center"><b><u><a href="https://www.snscart.com/admin/">Zip2Tax</a></u></b><br>
			</td>
		</tr>
		<tr>
			<td align="center">
				Username: 3233</br>
				Password: get2it
			</td>
			<td align="center">
				Username: 17202</br>
				Password: get2it
			</td>
		</tr>
	</table> 

			<table cellpadding="10" width="96%"> 
				<tr>
					<td>				
							<hr>
						<span class="newStyleHeading">Phase One:</span> <b>Receive Order from Shopping Cart - Initiated by a shopping cart sending us an order</b>
						
		
						<ul>
							<li>Accepts the request variable from carts</li>
							<li>Writes data to Barley1.ha_prod.ni_UploadFromShoppingCart and Casper08.ha_prod.ha_UploadFromShoppingCart
							<li>Sends telephone texts</li>
							<li>Sends Google conversion notification</li>
							<li>Sends Yotpo Notification</li>
						</ul>
						
						<hr>
						
						<span class="newStyleHeading">Phase Two:</span> <b>Create Order - Initiated by Barley.ha_prod.ni_UploadFromShoppingCart</b><br>
						
						<br>1.) exec ni_ShoppingCart_Add_DateCode<br>
						<i>- Posts a dupe code to each new record</i><br>
						<br>2.) exec ni_ShoppingCart_Add_Account<br>
						<i>- Posts the order including a new account if necessary</i><br>
						
						<hr>
						
						<span class="newStyleHeading">Phase Three:</span> <b>Fulfill the Requirements - Initiated by Casper10.ha_processes.dbo.ha_Processes_min,<br>
						Barley.ha_prod.dbo.ha_Requirements_check</b><br>
						<br><span class="subHead">Sets requirement flags in the order table based on what it finds in the order details table.</span><br>

						Barley.ha_prod.dbo.ni_Orders_Registration_Check<br>

						Barley.ha_prod.dbo.ni_Orders_Registration_Email<br><br>

						<span class="subHead">New Requirements System</span><br>
						Barley.ha_prod.dbo.ha_Requirement_z2t_lookup<br>
						Barley.ha_prod.dbo.ha_Requirement_z2t_lookup_Email<br>
						Barley.ha_prod.dbo.ha_Requirement_z2t_link<br>
						Barley.ha_prod.dbo.ha_Requirement_z2t_initial<br>
						Barley.ha_prod.dbo.ha_Requirement_shipping
					</td>
				</tr>   
			</table>		

		
		
		</td>
	</tr>
  </table>
  <!--#include virtual="ha_BackOffice/includes/BodyParts/ha_PageTail.inc"-->
</body>
</html>

