<!DOCTYPE html>
<!--#include virtual="includes/connection.asp"-->
<!--#include virtual="ha_backoffice/includes/Config.asp"-->
<!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->

<%
	Session("Redirect") = ""
	ColorTab = 7
	PageHeading = "Constant Contact Mailing Lists Export Utility"


    Session("CurrentPage") = "ha_ConstantContact_optimized.asp"
  
   Dim outfile
   
	SELECT Case Request("DownloadAction")
	Case "RaffleTicketCustomers"
		SQL =  "ni_emails_softwareandpaper"
		outfile = "rt_email_list.csv"
	Case "AllZip2TaxCustomers"
		SQL = "z2t_emails_list"
		outfile = "z2t_email_list.csv"
	Case "TableUpdateCustomers"
		SQL = "z2t_emails_list_existing_tables_customers_Optimized" '
		outfile = "z2t-table-update-customers.csv"
	Case "FreeTrialCustomersComplete"
		SQL="z2t_Free_Trial_Signups_Complete"
	Case "FreeTrialCustomersIncomplete"
		SQL="z2t_Free_Trial_Signups_Incomplete"
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

	<script language="javascript" type = "text/javascript">
				var foc=0;		
				
				// AJAX Handler
		function XHConn()
		{
		  openModal();
		
		  var xmlhttp, bComplete = false;
		  try { xmlhttp = new ActiveXObject("Msxml2.XMLHTTP"); }
		  catch (e) { try { xmlhttp = new ActiveXObject("Microsoft.XMLHTTP"); }
		  catch (e) { try { xmlhttp = new XMLHttpRequest(); }
		  catch (e) { xmlhttp = false; }}}
		  if (!xmlhttp) return null;
		  this.connect = function(sURL, sMethod, sVars, fnDone)
		  {
		
			if (!xmlhttp) return false;
			bComplete = false;
			sMethod = sMethod.toUpperCase();
			try {
			  if (sMethod == "GET")
			  {
				xmlhttp.open(sMethod, sURL+"?"+sVars, true);
				sVars = "";
			  }
			  else
			  {
				xmlhttp.open(sMethod, sURL, true);
				xmlhttp.setRequestHeader("Method", "POST "+sURL+" HTTP/1.1");
				xmlhttp.setRequestHeader("Content-Type",
				  "application/x-www-form-urlencoded");
			  }
			  xmlhttp.onreadystatechange = function(){
				if (xmlhttp.readyState == 4 && !bComplete)
				{
				  bComplete = true;
				  fnDone(xmlhttp);	
				  closeModal();
				}};
			  xmlhttp.send(sVars);
			}
			catch(z) { return false; }
			return true;
		  };
		  return this;
		}
		
		
		var doAJAXCall = function (PageURL, ReqType, PostStr, FunctionName) {
			// create the new object for doing the XMLHTTP Request
			var myConn = new XHConn();
			// check if the browser supports it
			if (myConn)	{
				// XMLHTTPRequest is supported by the browser, continue with the request
				myConn.connect('' + PageURL + '', '' + ReqType + '', '' + PostStr + '', FunctionName);    
			} 
			else {
				// Not support by this browser, alert the user
				alert("XMLHTTP not available. Try a newer/better browser, this application will not work!");   
			}
		}
		
		function openModal() {
			$('#modal').css("display",'block');
			$('#fade').css("display", 'block');
		}
		
		function closeModal() {
			$('#modal').css("display",'none');
			$('#fade').css("display", 'none');
		}

		function clickPopup(ID)
		{
			var url="/ha_Backoffice/Company/DocumentMaintenance/ha_Document_Show.asp";
			var param='PhotoID=' + ID;
			doAJAXCall(url,'GET',param ,ShowDocument);
		}
		
		function ShowDocument(oXML){
			
			$("#modal-report").html(oXML.response);
			
			$('#fade-report').css("display", 'block');
			$("#modal-report").css("display", 'block');//.show();
		
			$(".closeReport").click( function(){
			$('#fade-report').css("display", 'none');
			$("#modal-report").css("display","none");});

		}
		 			
		function clickDownload(d)
			{
			var URL = '/ha_backoffice/Utilities/ha_ConstantContact_optimized.asp' +
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

	#fade {display: none;position:absolute; top: 0%; left: 0%; width: 100%;height: 125%;
			 background-color: #ababab; z-index: 1001;-moz-opacity: 0.8;opacity: .70;filter: alpha(opacity=80);}
	#modal {display: none;position: absolute;top: 45%;left: 45%;width: 64px;height: 64px; padding:30px 15px 0px;
			border: 3px solid #ababab;box-shadow:1px 1px 10px #ababab;border-radius:20px;background-color: white;
			z-index: 1002;text-align:center;overflow: auto;}
	.closeReport {background: #606061; color: #FFFFFF;	line-height: 25px;	position: absolute;	right: -12px;	
			text-align: center;	top: -10px;	width: 24px;	text-decoration: none;	font-weight: bold;	
			-webkit-border-radius: 12px;-moz-border-radius: 12px;border-radius: 12px;
			-moz-box-shadow: 1px 1px 3px #000;-webkit-box-shadow: 1px 1px 3px #000;	box-shadow: 1px 1px 3px #000;}
	.closeReport:hover { background: #00d9ff; }
	#modal-report {display: none;position: absolute;top: 5%;left: 15%;width: 900px;
					height: 550px; padding:30px 15px 0px;border: 3px solid #ababab;box-shadow:1px 1px 10px #ababab;
					border-radius:20px;background-color: white;	z-index:1002;text-align:center;}
	.print{background-image: url(../images/download.jpg);}
	#fade-report {display: none;position:absolute; top: 0%; left: 0%; width: 100%;height: 125%;
		 background-color: #ababab; z-index: 1001;-moz-opacity: 0.8;opacity: .70;filter: alpha(opacity=80);}			
				
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
 <div id="fade"> </div>

 <div id="modal">
     <img id="loader" src="../../images/AjaxLoading.gif" />
     <div id="modal-content"></div>
 </div>
 <div id="modal-report">
 </div>
 <div id="fade-report"> </div>
 <div id="modal-report">
     <div id="modal-content">
     
     </div>
 </div>

  </body>
</html>

