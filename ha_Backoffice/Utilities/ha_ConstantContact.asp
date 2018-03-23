<!DOCTYPE html>
<!--#include virtual="includes/connection.asp"-->
<!--#include virtual="ha_backoffice/includes/Config.asp"-->
<!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->

<%
	Session("Redirect") = ""
	ColorTab = 7
	PageHeading = "Constant Contact Mailing Lists Export Utility"


    Session("CurrentPage") = "ha_ConstantContact.asp"
  
   Dim outfile
   
	SELECT Case Request("DownloadAction")
	Case "BothSoftAndPaper"
		SQL =  "ha_Customer_Emails_rts_paper_list"
		outfile = "ha-Customer-Emails-rts-paper_list.csv"
	Case "SoftwareRT"
		SQL = "ha_Customer_Emails_by_ProductCategory 'Software-RT'"
		outfile = "ha_email_list_Software.csv"
	Case "Paper"
		SQL = "ha_Customer_Emails_by_ProductCategory 'Paper'"
		outfile = "z2t_email_list_paper.csv"
	Case "TableUpdateCustomers"
		SQL = "z2t_existing_tables_customers_List	" '
		outfile = "z2t-table-update-customers.csv"
	Case "FreeTrialCustomersComplete"
		SQL="z2t_Free_Trial_Signups_Complete"
		outfile = "z2t-Free-Trial-Signups-Complete.csv"
	Case "FreeTrialCustomersIncomplete"	
		SQL="z2t_Free_Trial_Signups_Incomplete"
		outfile = "z2t-Free-Trial-Signups-Incomplete.csv"
	Case "CustomerRenewals"
		SQL="z2t_Customer_Renewals_list"
		outfile = "z2t-Customer-Renewals.csv"
	CASE "AllCustomers"
		SQL="z2t_Customer_Emails_list"
		outfile = "z2t-all-Customers-List.csv"
	
	Case Else
		SQL = ""
	End Select

	If SQL > "" Then
		'Do a download
	   Dim rs
	   set rs=server.createObject("ADODB.Recordset")
	   Dim connBackOffice
	   set connBackOffice=server.CreateObject("ADODB.Connection")
	   connBackOffice.Open "driver=SQL Server; server=66.119.50.228,7843; uid=davewj2o; pwd=get2it; database=ha_CRM"  'Casper08
	
	   rs.open SQL, connBackOffice, 2, 3

	   If Not rs.eof Then
		 ' Download the email list
		 Response.Clear
		 Response.AddHeader "Content-Disposition", _
							"attachment; filename=" _
							& outfile
		 Response.ContentType = "application/octet-stream"
		 Response.CharSet = "UTF-8"

		 Dim fieldcount

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
	   Response.End
	   Else
		'''do nothing
		End If
	End If
%>

<html>
<head>
    <title>Harvest American Backoffice -  <%=PageHeading%></title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

    <link href="<%=strPathIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet">
    <link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>menu/dropdowntabfiles/halfmoontabs.css">
		
    <script type="text/javascript" src="https://ajax.googleapis.com/ajax/libs/jquery/1.3.2/jquery.min.js"></script>
    <script type="text/javascript" src="<%=strPathIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>
	<script type="text/javascript" src="<%=strPathBOIncludes%>haFunctions.js"></script>
  	<script type="text/javascript" src="<%=strPathIncludes%>AJAXHandler.js"></script>

	<script language="javascript" type = "text/javascript">
	$(document).ready(function(e) {
              CountRecords();
    });
	
    function CountRecords() {
        
        // build up the post string when passing variables to the server side page
        var PostStr = "";
        
        // use the generic function to make the request
        doAJAXCall('ajax/ha_ResultSet_ajax.asp', 'POST', '' + PostStr + '', ShowCountOnPage );
    

    }
	
	function ShowCountOnPage(oXML){
		var response = oXML.response;
		var arrExportTypeSpans = ["ha_Software","ha_Paper","ha_Both","z2t_TableUploadCustomer","z2t_FreeTrialIncomplete","z2t_FreeTrialComplete", "z2t_CustomerRenewals","z2t_AllCustomers"];
		var	arrCount=	response.split("~");

		for (i = 0; i<arrCount.length;i++){
			 document.getElementById(arrExportTypeSpans[i]).innerHTML= numberWithCommas(arrCount[i]);
		}				
	}
	function numberWithCommas(x) {
    return x.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}
	
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
			#loader{
				width:16px;
				}
	  </style>
		
</head>

<body class="gray_desktop">
  <table align="center" cellspacing="0" cellpadding="0" class="MainBodyFixed">
	<tr valign="top">
	  <td>
		 <!--#include virtual="ha_Backoffice/includes/bodyparts/ha_PageHeading.inc"-->
         <!--#include virtual="ha_Backoffice/utilities/ConstantContact/ha_ContactConstantBody.inc"-->
      </td>
	</tr>
  </table>

    <!--#include virtual="ha_BackOffice/includes/BodyParts/ha_PageTail.inc"-->

  </body>
</html>

