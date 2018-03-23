<!DOCTYPE html>
<!--#include virtual="includes/connection.asp"-->
<!--#include virtual="ha_backoffice/includes/Config.asp"-->
<!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->

<%
	Session("Redirect") = ""
	ColorTab = 7
	PageHeading = "Constant Contact Mailing Lists Export Utility"


    Session("CurrentPage") = "boConstantContact.asp"
  
   ' Upload the file
   Dim infile
   filename = "c:\Zip2Tax\Export\email\bounces.csv"
   'chunk = Request.BinaryRead(200000)
   'Response.Write(chunk)

   z2tStatus = ""
   rtsStatus = ""

   Dim z2tStatus
   Dim rtsStatus

   Dim outfile
   
	
	SQL = "ha_ConstantContact_emailList 'c:\Zip2Tax\Export\email\bounces.csv', "
	SELECT Case Request("DownloadAction")
	Case "RaffleTicketCustomers"
		SQL = SQL & "'rts'"
		rtsStatus = "No bounced emails"
		outfile = "rt_email_list.csv"
	Case "AllZip2TaxCustomers"
		SQL = SQL & "'z2t'"
		z2tStatus = "No bounced emails"
		outfile = "z2t_email_list.csv"
	Case "TableUpdateCustomers"
		SQL = "exec('ha_prod.dbo.z2t_emails_list_existing_tables_customers')" '
		z2tTableUpdateStatus = "No table update customers"
		outfile = "z2t-table-update-customers.csv"
	Case Else
		SQL = ""
	End Select
	
	If SQL > "" Then
		'Do a download
	   Dim rs
	   set rs=server.createObject("ADODB.Recordset")
	   Dim connBackOffice
	   set connBackOffice=server.CreateObject("ADODB.Connection")
	   connBackOffice.Open "driver=SQL Server; server=66.119.50.228,7843; uid=davewj2o; pwd=get2it; database=ha_Prod"  'Casper08
	   '"driver=SQL Server; server=casper10.harvestamerican.net,7043; uid=davewj2o; pwd=get2it; database=ha_BackOffice"
	   Response.Write(SQL)
	   rs.open SQL, connBackOffice, 2, 3

	   Dim status
	   status = ""

	   If Not rs.eof Then
		 ' Download the email list
		 Response.Clear
		 Response.AddHeader "Content-Disposition", _
							"attachment; filename=" _
							& outfile
		 Response.ContentType = "application/octet-stream"
		 Response.CharSet = "UTF-8"

		 Dim fieldcount

		 ' column headings
	'     fieldcount = 0
	'     For Each field in rs.Fields
	'       fieldcount = fieldcount + 1
	'       Response.Write(field.name)
	'       If fieldcount < rs.Fields.Count Then Response.Write(",")
	'     Next
	'     Response.Write(chr(13) + chr(10))

		 Dim count
		 count = 0
		 Do While Not rs.eof
		   count = count + 1
		   fieldcount = 0
		   For Each field in rs.Fields
			 fieldcount = fieldcount + 1
			 Response.Write(field)
			 If fieldcount < rs.Fields.Count Then Response.Write(",")
		   Next
		   Response.Write(chr(13) + chr(10))
		   rs.moveNext
		 Loop
		 Response.End
		 rs.close
		 Set rs = Nothing
		 If Request("z2t-submit") Then
		   z2tStatus = cStr(count) & " bounced emails"
		 Else
		   rtsStatus = cStr(count) & " bounced emails"
		 End If
	   Response.End
	   Else
		 status = "No bounced emails"
		End If
	End If
%>
 
<html>
	<head>
		<title>Harvest American Backoffice -  <%=PageHeading%></title>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

		<link href="<%=strPathIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet">
		<link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>menu/dropdowntabfiles/halfmoontabs.css">
		
		<script type="text/javascript" src="<%=strPathBOIncludes%>haFunctions.js"></script>
		<script language="javascript" type = "text/javascript">
			function clickPopup(ID)
				{
				var URL = '/ha_Backoffice/Company/DocumentMaintenance/ha_Document_Show.asp' +
					'?PhotoID=' + ID;
					void window.open(URL,'900','500')
				}
		 			
			function clickDownload(d)
				{
				var URL = '/ha_backoffice/Utilities/ha_ConstantContact.asp' +
					'?DownloadAction=' + d;
					window.location = URL;
				}
		</script>
		
		<script type="text/javascript" src="<%=strPathIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>
	 
		<style type="text/css">
	  
			span.subHead
				{
				font-weight: 	bold;
				color:			#336600;
				}
			span.superHead
				{
				font-size:		18pt;
				font-weight:	bold;
				
				}
			.titles
				{
				padding-left:3em;
				}
				
				
			.newStyleHeading
				{
				font-size: 		11pt;
				font-weight: 	bold;
				color:			#336600;
				}
				
	  </style>
		
</head>

<body class="gray_desktop">
  <table align="center" cellspacing="0" cellpadding="0" class="MainBodyFixed">
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
					Constant Contact is a vendor who sends bulk E-mails for us.  Our mailing lists with them
					needs to be frequently updated.  Here you can export lists from our data into csv files 
					for uploading into Constant Contact.
				  </td>
				</tr>

				<tr><td>&nbsp;</td></tr>
				
				<tr>
				  <td align="right">
					<a href="javascript:clickPopupImage(1405);" class="TopRightLink">How to Work with Constant Contact Mailing Lists</a>
				  </td>
				</tr>
		  
				<tr><td>&nbsp;</td></tr>
				
				<tr>
				  <td>
					<table width="100%" cellpadding="0" cellspacing="0">
					  <tr>
					    <td width="85%">
						  <span class="newStyleHeading">Harvest</span> (Constant Contact account davewj2o)<br><br>
										
						  This exports the account e-mail address and first name of every <b>Raffle Ticket Software</b> customer
						  who has ever purchased ticket paper or the Raffle Ticket Software.<br>
						</td>
						<td align="right" valign="bottom" width="15%">
						  <a href="javascript:clickDownload('RaffleTicketCustomers');" class="greenButton">Download</a>
						</td>
					  </tr>
					  
					  <tr><td>&nbsp;</td></tr>
					  </tr><td colspan="2"><hr></td></tr>
					  <tr><td>&nbsp;</td></tr>
									
					  <tr>
					    <td>
						  <span class="newStyleHeading">Zip2Tax</span> (Constant Contact account info@Zip2Tax.com)<br><br>
										
						  These contacts are for the list <b>Zip2Tax Customers</b>. 
						  It exports the account e-mail address and account first name of every Zip2Tax customer who has
						  ever purchased any Zip2Tax product.										
						</td>
						<td align="right" valign="bottom">
						  <a href="javascript:clickDownload('AllZip2TaxCustomers');" class="greenButton">Download</a>
						</td>
					  </tr>
					  
					  <tr><td>&nbsp;</td></tr>
					  
					  <tr>
					    <td>
						  These contacts are for the list <b>Zip2Tax Table Update Customers</b>. 
						  This exports the account's first name along with the "Product Edit" screen's e-mail address(es), 
						  the login, and the password, for all current Zip2Tax customers.<br><br>
						  <i>Note: delete the existing "Zip2Tax Table Update Subscribers" mailing list prior to
						  importing this new list to remove expired subscriptions.</i>
						</td>
						<td align="right" valign="bottom">
						  <a href="javascript:clickDownload('TableUpdateCustomers');" class="greenButton">Download</a>
						</td>
					  </tr>
					  
					  <tr><td>&nbsp;</td></tr>
					  </tr><td colspan="2"><hr></td></tr>
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

