<%
	Session("Redirect") = ""
	ColorTab = 1
	PageHeading = "Table Locations (Published Tables)"
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
			
		td.farm
			{
			width:			25%;
			border:			1px solid black;
			padding:		5px;
			vertical-align:	top;
			background-color: #FFFFCC;
			}
			
		td.farm span.head
			{
			font-weight:	bold;
			font-size:		11pt;
			text-align:		center;
			border-bottom:	1px solid black;
			width:			96%;
			display:		block;
			margin:			0 auto 0 auto;
			}
			
		td.farm ul
			{
			margin:			0 0 0 20;
			padding:		2 0 2 0;
			}
			
		span.obsolete
			{
			font-family: 	Century;
			font-style: 	italic;
			}
			
		li
			{
			font-size:		9pt;
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
		
		<table width="98%" border="0" cellspacing="3" cellpadding="0" align="center">
	
		  <tr><td>&nbsp;</td></tr>
		
		  <tr>
			<td class="farm">
			  <span class="head">Barleys</span>
			  <ul>
			    <li><b>Barley</b></li>
				<ul>
				  <li>ha_prod</li>
				  <ul>
  					<li><span class="obsolete">ni_Accounts</span></li>
					<li><span class="obsolete">ni_Orders</span></li>
					<li><span class="obsolete">ni_OrderDetails</span></li>
				  </ul>
				</ul>
			  </ul>
			  <ul>
			    <li><b>Barley2</b></li>
				<ul>
				  <li>z2t_WebPublic</li>
				  <ul>
					<li><span class="obsolete">z2t_Accounts_Subscriptions</span></li>
					<li><span class="obsolete">z2t_Accounts_Subscriptions_addons</span></li>
				  </ul>
				</ul>
			  </ul>
			</td>
			
			<td class="farm">
			  <span class="head">Caspers</span>
			  <ul>
			    <li><b>Casper06</b></li>
			  </ul>
			  <ul>
			    <li><b>Casper07</b></li>
			  </ul>
			  <ul>
			    <li><b>Casper08</b></li>
				<ul>
				  <li>ha_CRM</li>
				  <ul>
					<li>ha_Orders</li>
					<li>ha_OrderDetails</li>
				  </ul>
				</ul>
			  </ul>
			  <ul>
			    <li><b>Casper09</b></li>
			  </ul>
			  <ul>
				<li><b>Casper10</b></li>
				<ul>
				  <li>ha_PublishedTables</li>
				  <ul>
					<li>ha_Accounts</li>
					<li>ha_Products</li>
					<li>ha_Types</li>
				  </ul>
				</ul>
				<ul>
				  <li>z2t_BackOffice</li>
				  <ul>
					<li>z2t_Activity_20??</li>
					<li>z2t_TableDownload_Log</li>
				  </ul>
				</ul>
			  </ul>
			</td>
			
			<td class="farm">
			  <span class="head">Phillys</span>
			  <ul>
			    <li><b>Philly01</b></li>
				<ul>
				  <li>z2t_PublishedTables</li>
				  <ul>
					<li>z2t_Accounts_Subscriptions</li>
					<li>z2t_Accounts_Subscriptions_addons</li>
					<li>z2t_EffectiveDates</li>
					<li>z2t_Special_Rules</li>
					<li>z2t_StateInfo</li>
					<li>z2t_Types</li>
				  </ul>
				</ul>
			  </ul>
			  <ul>
			    <li><b>Philly02</b></li>
			  </ul>
			  <ul>
			    <li><b>Philly03</b></li>
			  </ul>
			  <ul>
			    <li><b>Philly04</b></li>
			  </ul>
			  <ul>
			    <li><b>Philly05</b></li>
				<ul>
				  <li>z2t_BackOffice</li>
				  <ul>
					<li>z2t_Activity_20??</li>
					<li>z2t_TableDownload_Log</li>
				  </ul>
				</ul>
				<ul>
				  <li>z2t_UpdateRates</li>
				  <ul>
					<li>z2t_TaxData</li>
				  </ul>
				</ul>
			  </ul>
			</td>
			
			<td class="farm">
			  <span class="head">Franks</span>
			  <ul>
			    <li><b>Frank01</b></li>
			  </ul>
			  <ul>
			    <li><b>Frank02</b></li>
			  </ul>
			  <ul>
			    <li><b>Frank03</b></li>
			  </ul>
			</td>
		  </tr>
		  
		  <tr><td>&nbsp;</td></tr>
		  
		  <tr>
			<td colspan="4">
			  Table names in <span class="obsolete">italics</span> = soon to be obsolete.
			</td>
		  </tr>
		  
		  <tr>
			<td colspan="4">
			  All other tables using these same names are probably replicas.
			</td>
		  </tr>
		  
  		  <tr><td>&nbsp;</td></tr>

		</table>

	  </td>
	</tr>
  </table>
  <!--#include virtual="ha_BackOffice/includes/BodyParts/ha_PageTail.inc"-->
</body>
</html>

