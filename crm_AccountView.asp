<!--#include virtual="includes/BodyParts/PageStart.asp"-->
<!--#include virtual="includes/crm_connection.asp"-->
<!--#include virtual="includes/crm_connection_PublishedTables.asp"-->
<!--#include virtual="includes/old/functions.asp"-->
<!--#include virtual="includes/old/ha/connection.asp"-->
<!--#include virtual="includes/old/ha/ha/ha_functions.asp"-->

<%
    dim strColor
    dim LineCount
    dim dateNow

    dateNow = Escape(FormatMediumDate(now))

    set RS=server.createObject("ADODB.Recordset")

    strHeading = "Account View"
    Session("CurrentPage") = "crmAccountView.asp"
    Session("Section")="Accounts"
    Session("aMode")="null"

    If Session("oMode")="Add" Then
	'We just added an order . . . go to it.
        response.redirect strPathBO & "boOrderView.asp?OrderID=" & Session("OrderID")
    End If

    If Request("ID")="" or isnull(Request("ID")) then
		If Session("AccountID")="" or isnull(Session("AccountID")) then
			Response.Redirect strPathBO & "crm_Accounts.asp"
		End If
    Else
		Session("AccountID") = Request("ID")
		AccountID = Request("ID")
    End If



    If Request("aMode")="RegSync" Then
        objcon.execute "ni_Registrations_sync"
		Response.Redirect "crm_AccountView.asp"
    End if

    If Request("aMode")="Del" Then
        objcon.execute "ni_Order_delete(" & Request("OrderID") & ", '" & Session("Login") & "')"
		Response.Redirect "crm_AccountView.asp"
    End if

    If Request("aMode")="z2tDel" Then
        objcon.execute "z2t_Subscription_Delete(" & Request("z2tID") & ", '" & Session("Login") & "')"
		Response.Redirect "crm_AccountView.asp"
    End if

    If Request("aMode")="Duplicate" Then
        objcon.execute "ni_Order_duplicate(" & Request("OrderID") & ", '" & Session("Login") & "')"
		Response.Redirect "crm_AccountView.asp"
    End if

    if Request("cpMode")="Disable" Then
		SQL="SELECT [CPID], [DeletedBy], [DeletedDate] " & _
			"FROM [ni_CustomerPages] " & _
			"WHERE [CPID] = " & Request("CPID") 	
		rs.open SQL,objcon,1,3
		rs("DeletedBy") = Session("Login")
		rs("DeletedDate") = now
		rs.update
		rs.close
		Response.Redirect "crm_AccountView.asp"
    End if

    if Request("cpMode")="Enable" Then
		SQL="SELECT [CPID], [DeletedBy], [DeletedDate] " & _
			"FROM [ni_CustomerPages] " & _
			"WHERE [CPID] = " & Request("CPID") 	
		rs.open SQL,objcon,1,3
		rs("DeletedBy") = null
		rs("DeletedDate") = null
		rs.update
		rs.close
		Response.Redirect "crm_AccountView.asp"
    End if	

%>

<html>
<head>
    <title>Harvest American CRM Account View</title>
    
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	
    <link type="text/css" rel="stylesheet" href="http://crm.HarvestAmerican.info/includes/crm.css">

    <script type="text/javascript" src="http://crm.HarvestAmerican.info/includes/ha_functions.js"></script>
    <script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.3.2/jquery.min.js"></script>

	<style type="text/css">   
	#modal-email {display: none;position: absolute;top: 5%;left: 15%;width: 880px;
					height: 630px; padding:30px 15px 0px;border: 3px solid #ababab;box-shadow:1px 1px 10px #ababab;
					border-radius:20px;background-color: white;	z-index:1002;text-align:center;}
	#modal-Page {display: none;position: absolute;top: 5%;left: 15%;width: 400px;
					height: 200px; padding:30px 15px 0px;border: 3px solid #ababab;box-shadow:1px 1px 10px #ababab;
					border-radius:20px;background-color: white;	z-index:1002;text-align:center;}

	#fade {display: none;position:fixed; top: 0px; left: 0px; width: 100%;height: 100%;
			 background-color: #ababab; z-index: 1001;-moz-opacity: 0.8;opacity: .70;filter: alpha(opacity=80);}
	#modal {display: none;position: absolute;top: 45%;left: 45%;width: 64px;height: 64px; padding:30px 15px 0px;
			border: 3px solid #ababab;box-shadow:1px 1px 10px #ababab;border-radius:20px;background-color: white;
			z-index: 1002;text-align:center;overflow: auto;}
	.closeReport {background: #606061; color: #FFFFFF;	line-height: 25px;	position: absolute;	right: -12px;	
			text-align: center;	top: -10px;	width: 24px;	text-decoration: none;	font-weight: bold;	
			-webkit-border-radius: 12px;-moz-border-radius: 12px;border-radius: 12px;
			-moz-box-shadow: 1px 1px 3px #000;-webkit-box-shadow: 1px 1px 3px #000;	box-shadow: 1px 1px 3px #000;}
	.closeReport:hover { background: #00d9ff; }
	.print{background-image: url(../../images/download.jpg);}
	#fade-report {display: none;position:fixed; top: 0px; left: 0px; width: 100%;height: 100%;
		 background-color: #ababa b; z-index: 1001;-moz-opacity: 0.8;opacity: .70;filter: alpha(opacity=80);}			
    
      
             	</style>
   
<script language='JavaScript'>

	var AccountID = <%=Session("AccountID")%>	
	var updatedSection = '';
	
	window.onload = function() 
		{     
		//I don't remember what this is for
		if ("<%=Request("popup")%>" != "") window.resizeTo(700,600);

		//Refreshes only the section required upon regaining focus
		window.onfocus = function() 
			{ 
			if (updatedSection == 'Zip2Tax Products Ordered')
				{
				updatedSection = '';
				runZip2TaxInfo();
				}
				
			if (updatedSection != '')
				{
				location.reload();
				}
			}
		}
	 
    function clickTicketEdit(id)
        {
          if (id > '0')
          {
           var URL = 'http://helpdesk.harvestamerican.net/bouser/details.asp' +
            '?id=' + id;
               }
             else 
          {
           var URL = 'http://helpdesk.harvestamerican.net/bouser/new.asp' +
            '?acctid=<%=Session("AccountID")%>';
          }
        openPopUp(URL);
        }

    function clickEmailEditor(ID, EmailName)
        {
        var URL = '/ha_Backoffice/Email/ha_Email_Editor.asp' +
            '?ID=' + ID +
            '&e=' + EmailName;
        openPopUp(URL);
        }	
		

		//////////////////////////////
		
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
var doc = document.documentElement;
var top = (window.pageYOffset || doc.scrollTop)  - (doc.clientTop || 0);

	$('#modal').css("top",top);
    $('#modal').css("display",'block');
    $('#fade').css("display", 'block');
}

function closeModal() {
    $('#modal').css("display",'none');
    $('#fade').css("display", 'none');
}

function ShowPage(Url){
var url =	Url.split("?")[0];
var param=Url.split("?")[1];
	doAJAXCall(url,'GET',param,ShowPageResponse);
}
function SendInvoicePage(Url){
	var url =	Url.split("?")[0];
	var param=Url.split("?")[1];
	doAJAXCall(url,'GET',param,ShowPageResponse);
	
}

function LoadEMail(ID, EmailName){

	var url='http://crm.HarvestAmerican.info/accounts/Email/ha_Email_Editor.asp';
	var param='ID=' + ID +'&e=' + EmailName; 

	doAJAXCall(url,'GET',param ,EmailForProducts);
}
function sendEmail( ){

	var eSendFrom =$("#SendFrom").val();
	 var eSendTo =$("#SendTo").val();
	var eCopy =$("#Copy").val();
	var eSubject =$("#Subject").val();
	var eText=$("#noteTextArea").text();
		var eSQL =$("#hdnSQL").val();
		
	var url='http://crm.HarvestAmerican.info/accounts/Email/ha_Email_Editor.asp';

	var param='eSQL='+eSQL+'&eSendFrom='+eSendFrom+'&eSendTo='+eSendTo+'&eCopy='+eCopy+'&eSubject='+eSubject+'&eText='+eText;
	//alert(param);
	doAJAXCall(url,'GET',param ,EmailResponse);
	
	}
	
function EmailResponse(oXML){
	var response = oXML.response;
	alert(response);

	$("#modal-Email").hide();
	$("#fade-report").hide();

}
function EmailForProducts (XML){

	var response=XML.response;
//	alert(response);
	$("#modal-Email").html(response);
	var doc = document.documentElement;
	var top = (window.pageYOffset || doc.scrollTop)  - (doc.clientTop || 0);

	$("#modal-Email").css("top",top);	

	$("#modal-Email").show();
	$("#fade-report").show();
	
	$(".closeReport").click( function(){
		$('#fade-report').css("display", 'none');
		$("#modal-Email").css("display","none");
  });	
	
}
function ShowPageResponse(oXML)
{
	var response = oXML.response;
		
	$("#modal-Page").html(response);
	var doc = document.documentElement;
	var top = (window.pageYOffset || doc.scrollTop)  - (doc.clientTop || 0);

	$("#modal-Page").css("top",top);	

	$("#modal-Page").show();
	$("#fade-report").show();

	$(".closeReport").click( function(){
		$('#fade-report').css("display", 'none');
		$("#modal-Page").css("display","none");
  });	
}
	
function querySt(Key) {
    var url = window.location.href;
    KeysValues = url.split(/[\?&]+/);
    for (i = 0; i < KeysValues.length; i++) {
        KeyValue = KeysValues[i].split("=");
        if (KeyValue[0] == Key) {
            return KeyValue[1];
        }
    }
}
function printDiv(printpage)
{
	var headstr = "<html><head><title></title></head><body>";
	var footstr = "</body>";
	var oldstr = document.body.innerHTML;

// remove print and close buttons
	$("#"+printpage).find($("[id^=print]").remove());
	$("#"+printpage).find($("[class=closeReport]").remove());
	
	if(printpage=="modal-Email");
	{
		$("#"+printpage).find(".buttons").remove();	
	}
	
	var newstr = $("#"+printpage).html();
	
	document.body.innerHTML = headstr+newstr+footstr;

	window.print();
	document.body.innerHTML = oldstr;
	$(".closeReport").click( function(){
	$('#fade-report').css("display", 'none');
	$("#"+printpage).css("display","none");

	return false;
  });	
}



</script>

</head>

<body class="gray">

<table align="center" cellspacing="0" cellpadding="0" class="MainBody">

  <!--#include virtual="includes/BodyParts/BodyTop.inc"-->

  
<!--#include file="crm_AccountView_Information.inc"-->
<!--#include file="crm_AccountView_Orders.inc"-->
<!--#include file="crm_AccountView_SoftwareRegistrations.inc"-->
<!--#include file="crm_AccountView_PersonalPages.inc"-->
<!--#include file="crm_AccountView_Zip2Tax.inc"-->
<!--#include file="crm_AccountView_Activity.inc"-->
<!--#include file="crm_AccountView_Notes.inc"-->
  
</table>

<script language='JavaScript'>
	//clip.glue('d_clip_button');
	//clip.glue( 'd_clip_button', 'd_clip_container' );
</script>

</body>

<%
    Function subHead(s, w)

        subHead = "<th width='" & w & "%' class='sectionSubHead'>" & s & "</th>"

    End Function

Function FormatMediumDate(DateValue)
    Dim strYYYY
    Dim strMM
    Dim strDD

        strYYYY = CStr(DatePart("yyyy", DateValue))

        strMM = CStr(DatePart("m", DateValue))
        If Len(strMM) = 1 Then strMM = "0" & strMM

        strDD = CStr(DatePart("d", DateValue))
        If Len(strDD) = 1 Then strDD = "0" & strDD

        FormatMediumDate = strMM & "/" & strDD & "/" & strYYYY

End Function 

Sub NotesRead(AccountID, NoteID, BoxHeight)

    set rs=Server.CreateObject("ADODB.Recordset")
    strSQL = "ha_Notes_read(" & AccountID & ", " & NoteID & ", " & BoxHeight & ", -1)"

    rs.open strSQL, conn, 3, 3, 4

    if not rs.eof then
        do while not rs.eof
            Response.write rs("Result") & chr(10)
            rs.movenext
        loop
    end if

    rs.close
    set rs = nothing

End Sub

%>
 <div id="fade"> </div>

 <div id="modal">
     <img id="loader" src="../../images/AjaxLoading.gif" />
     <div id="modal-content"></div>
 </div>
<div id="modal-Page"></div>
 <div id="fade-report"> </div>

 <div id="modal-Email">
     <div id="modal-email-content">
	 	<a href="#" title="Close" class="closeReport">X</a>     
     </div>
 </div>
	
</html>
