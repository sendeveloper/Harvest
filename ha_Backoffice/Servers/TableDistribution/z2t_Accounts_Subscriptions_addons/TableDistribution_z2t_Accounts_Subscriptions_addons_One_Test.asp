<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

<!--#include virtual="ha_backoffice/includes/Config.asp"-->
<!--#include virtual="ha_backoffice/Servers/TableDistribution/includes/TableDistribution_ReadObjects.inc"-->

<%
	If isnull(Request("z2tID")) or Request("z2tID") = "" Then
		z2tID = ""
	Else
		z2tID = Request("z2tID")
	End If
	
	Call ReadObjects("z2t_Accounts_Subscriptions_addons")
%>

<html>
<head>
    <title>Table Distribution Status z2t_Accounts_Subscriptions_Addons</title>

    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>ha_BackofficePopup.css">

    <script language="javascript" type="text/javascript" src="<%=strPathBOIncludes%>haFunctions.js"></script>
		
<script language='JavaScript'>

	// var http = new XMLHttpRequest();
	var loopStep = 1;
	var z2tID = <%=z2tID%>;	
	var locationCount = <%=LocationCount%>;
	var locationServer = [ 
		<%For i = 0 to LocationCount
			Response.write "'" & LocationServer(i) & "', "
		Next%>];
	var locationDatabase = [ 
		<%For i = 0 to LocationCount
			Response.write "'" & LocationDatabase(i) & "', "
		Next%>];
	var updateFlag = 0;
	var totalInserted = 0;
	var totalUpdated = 0;
		
    function formLoad()
	{
        //SetScreen(950, 600);
		runUpdate();
	}
	
	function getURL(id)
	{
		var url = '<%=strPathServers%>TableDistribution/z2t_Accounts_Subscriptions_Addons/TableDistribution_z2t_Accounts_Subscriptions_addons_update.asp' +
				'?s=' + locationServer[id] +
				'&d=' + locationDatabase[id] +
				'&z2tID=' + z2tID +
				'&ExecBy=TableDistribution_z2t_Accounts_Subscriptions_addons_One.asp' + 
				'&Now=' + escape(Date());
		return url;
	}			
	function runUpdate() 
	{
		var http = [];

		var successFn1 = function(ignore, r, timer) 
        {
            clearTimeout(timer);
            getUpdateResponse(ignore, r, 1);
        };

        var errorFn1 = function(r, ignore, timer) 
		{
            clearTimeout(timer);
            errorfn(r, ignore, 1);
        };

        var timeoutFn1 = function(ignore, ignore2, timer) 
        {
            clearTimeout(timer);
            linkToError("ajax timeout", "Timed out with no response.")(document.querySelectorAll("#resultServer" + 1)[0].firstChild);
        }

        http[1] = xhr(getURL(1), successFn1, errorFn1, 15000, timeoutFn1);
        errorFn1.timer = http.timer;
        successFn1 = http.timer;
        http[1].get();
	}

    function getUpdateResponse(ignore, r, id) 
	{
        var eleID = 'resultServer'+id.toString();
		document.getElementById(eleID).innerHTML = r.getElementsByTagName('response_server_name')[0].firstChild.nodeValue;
		var eleID = 'resultDatabase'+id.toString();
		document.getElementById(eleID).innerHTML = r.getElementsByTagName('response_database_name')[0].firstChild.nodeValue;
		var eleID = 'resultUpdated'+id.toString();
		document.getElementById(eleID).innerHTML = r.getElementsByTagName('response_records_updated')[0].firstChild.nodeValue;
	}
	function linkToError(title, message) {
		return function showError(e) {
	        var parent = e.parentNode
	        parent.innerHTML = "<span class=\"error\">" + title + "</span>";
	        parent.style.cursor = "pointer";
	        parent.childNodes[0].style.cursor = "pointer";
	        parent.addEventListener("click", function(e){alert(title + ":\n\n" + message);}, false);
	   	}
   	}
	function undef(obj, /*optional*/ alternative) {
		var altform = !(typeof(alternative) == "undefined")
		var isUndefined = (typeof(obj) == "undefined");
		return altform ? (isUndefined ? alternative : obj) : isUndefined;
 	}
 	function xhr(url, /*optional*/ fn, error, timeout, timeoutfn) {

		var http = new XMLHttpRequest();
		http.get = function xhrGet(/*optional*/ asynchronous) {
		http.open("GET", url, undef(asynchronous, true));
		//http.onabort = timeoutfn;
		http.timer = setTimeout(function(){http.abort(); timeoutfn();}, timeout ? timeout : 8000);
		http.send()
		return http;};

		http.post = function xhrPost(params, /*optional*/ asynchronous) {
		http.open("POST", url, undef(asynchronous, true));
		//http.abort = timeoutfn;
		http.timer = setTimeout(function(){http.abort(); timeoutfn();}, timeout ? timeout : 8000);
		http.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
		http.setRequestHeader("Content-length", params.length);
		http.setRequestHeader("Connection", "close");
		http.send(params);
		return http;}

		http.onreadystatechange = xhrChange;
		http.fn = undef(fn, xhrDefaultFn);
		http.error = undef(error, xhrDefaultError);
		http.url = url;
		//http.send(null); // allow caller to send post data or store ancillay data in the http request obect
		return http;
  	}
  	function xhrChange(state) {
		try {
			var states = {0: "uninitialized", 1: "loading", 2: "loaded", 3: "interactive", 4: "complete"};
			var status = {0: "unreachable", 404: "not found", 200: "success", 500: "server error"};

			switch (this.readyState) {
				case 4:
			  		switch (this.status) {
			  			case 200:
			      			this.fn.call(this, this.responseText, this.responseXML, this.timer);
			    			break;
			  			default:
			    			this.abort();
			    			this.error.call(this, this.responseText, this.responseXML, this.timer);
			    			return;
			    			break;
			    	}
			  		break;
				default:
					break;
			}
			return;
		} catch(error) {
			alert(error); return;
		}
	}
	function xhrDefaultFn(body) {
		var element = document.createElement("textarea");
		document.body.appendChild(element);
		element.style.width = "100%";
		element.innerHTML = body;
		return;
	}
	function xhrDefaultError(body) {
		switch (this.status) {
			case 0:
				alert("Unreachable URL\n\n" + this.url);
				return;
				break;
			default:
				var errorText;
				alert("XHR error:\n\n" + body);
				return;
			break;
		}
		return;
	}
</script>

</head>

<body onLoad="formLoad();">

<table width="650" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td>

        <table width="100%" border="0" cellspacing="10" cellpadding="0" align="center">

          <tr>
            <td class="popupHeading">
              Update - z2t_Accounts_Subscriptions_Addons
            </td>
          </tr>

		  <tr>
		    <td>
		  
		<table width="100%" border="0" cellspacing="2" cellpadding="2">
				  
		  
		  <tr>
			<td width="15%">
			  &nbsp;
			</td>
			<td class="subHead" width="35%">
			  Server
			</td>
			<td class="subHead" width="35%">
			  Database
			</td>
			<td class="subHead" width="15%">
			  Count
			</td>
		  </tr>
		  
		  
<%
	For i = 1 to LocationCount
%>
		  
          <tr>
		    <td align="center">
              <%=i%>
            </td>
			<td align="center" id="resultServer<%=i%>">
			  - - - - -
			</td>
			<td align="center" id="resultDatabase<%=i%>">
			  - - - - -
			</td>
			<td align="center" id="resultUpdated<%=i%>">
			  -
			</td>
          </tr>
		  
<%
	Next
%>
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

<%
    'Post the page load into activity
    'dim connPublic
    'set connPublic=server.CreateObject("ADODB.Connection")
    'connPublic.Open "driver=SQL Server;server=dbWeb.Zip2Tax.com;uid=z2t_WebUser;pwd=WebUser_z2t;database=z2t_WebPublic"

    Data2 = ""
    Data1 = request("p")

    pgeURL = Request.ServerVariables("path_info")

    if left(pgeURL,1) = "/" then
        pgeURL = mid(pgeURL&"  ",2)
    end if

    pgeURL = trim(pgeURL)
    URL = Request.ServerVariables("HTTP_REFERER")

    SQL = "z2t_Activity_add_new('" & Session("z2t_UserName") & "', 18, " & _
            "'" & Data1 & "', " & _
            "'" & Data2 & "', " & _
            "'" & pgeURL & "', " & _
            "'" & URL & "', " & _
            "'" & Session.SessionID & "', " & _
            "'" & Request.ServerVariables("REMOTE_ADDR") & "', " & _
            "'z2t_SpecialRulesStatus.asp', " & _
            Session("CookieID") & ")"

    'connPublic.Execute(sql)
%>

</body>
</html>
