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
	// var updateFlag = 0;
	// var totalInserted = 0;
	// var totalUpdated = 0;
    // var countColumn = {};
    // var countState = {};
    // var displayedContent = false;
		
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

        var eleID = 'resultServer'+id.toString();
        document.getElementById(eleID).innerHTML = locationServer[id];
        eleID = 'resultDatabase'+id.toString();
        document.getElementById(eleID).innerHTML = locationDatabase[id];
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
            linkToError("ajax timeout", "Timed out with no response.")(document.querySelectorAll("#resultUpdated" + 1)[0].firstChild);
        }

        var successFn2 = function(ignore, r, timer) 
        {
            clearTimeout(timer);
            getUpdateResponse(ignore, r, 2);
        };

        var errorFn2 = function(r, ignore, timer) 
		{
            clearTimeout(timer);
            errorfn(r, ignore, 2);
        };

        var timeoutFn2 = function(ignore, ignore2, timer) 
        {
            clearTimeout(timer);
            linkToError("ajax timeout", "Timed out with no response.")(document.querySelectorAll("#resultUpdated" + 1)[0].firstChild);
        }

        var successFn3 = function(ignore, r, timer) 
        {
            clearTimeout(timer);
            getUpdateResponse(ignore, r, 3);
        };

        var errorFn3 = function(r, ignore, timer) 
		{
            clearTimeout(timer);
            errorfn(r, ignore, 3);
        };

        var timeoutFn3 = function(ignore, ignore2, timer) 
        {
            clearTimeout(timer);
            linkToError("ajax timeout", "Timed out with no response.")(document.querySelectorAll("#resultUpdated" + 3)[0].firstChild);
        }

        var successFn4 = function(ignore, r, timer) 
        {
            clearTimeout(timer);
            getUpdateResponse(ignore, r, 4);
        };

        var errorFn4 = function(r, ignore, timer) 
		{
            clearTimeout(timer);
            errorfn(r, ignore, 4);
        };

        var timeoutFn4 = function(ignore, ignore2, timer) 
        {
            clearTimeout(timer);
            linkToError("ajax timeout", "Timed out with no response.")(document.querySelectorAll("#resultUpdated" + 4)[0].firstChild);
        }

        var successFn5 = function(ignore, r, timer) 
        {
            clearTimeout(timer);
            getUpdateResponse(ignore, r, 5);
        };

        var errorFn5 = function(r, ignore, timer) 
		{
            clearTimeout(timer);
            errorfn(r, ignore, 5);
        };

        var timeoutFn5 = function(ignore, ignore2, timer) 
        {
            clearTimeout(timer);
            linkToError("ajax timeout", "Timed out with no response.")(document.querySelectorAll("#resultUpdated" + 5)[0].firstChild);
        }

        var successFn6 = function(ignore, r, timer) 
        {
            clearTimeout(timer);
            getUpdateResponse(ignore, r, 6);
        };

        var errorFn6 = function(r, ignore, timer) 
		{
            clearTimeout(timer);
            errorfn(r, ignore, 6);
        };

        var timeoutFn6 = function(ignore, ignore2, timer) 
        {
            clearTimeout(timer);
            linkToError("ajax timeout", "Timed out with no response.")(document.querySelectorAll("#resultUpdated" + 6)[0].firstChild);
        }

        var successFn7 = function(ignore, r, timer) 
        {
            clearTimeout(timer);
            getUpdateResponse(ignore, r, 7);
        };

        var errorFn7 = function(r, ignore, timer) 
		{
            clearTimeout(timer);
            errorfn(r, ignore, 7);
        };

        var timeoutFn7 = function(ignore, ignore2, timer) 
        {
            clearTimeout(timer);
            linkToError("ajax timeout", "Timed out with no response.")(document.querySelectorAll("#resultUpdated" + 7)[0].firstChild);
        }

        var successFn8 = function(ignore, r, timer) 
        {
            clearTimeout(timer);
            getUpdateResponse(ignore, r, 8);
        };

        var errorFn8 = function(r, ignore, timer) 
		{
            clearTimeout(timer);
            errorfn(r, ignore, 8);
        };

        var timeoutFn8 = function(ignore, ignore2, timer) 
        {
            clearTimeout(timer);
            linkToError("ajax timeout", "Timed out with no response.")(document.querySelectorAll("#resultUpdated" + 8)[0].firstChild);
        }

        var successFn9 = function(ignore, r, timer) 
        {
            clearTimeout(timer);
            getUpdateResponse(ignore, r, 9);
        };

        var errorFn9 = function(r, ignore, timer) 
		{
            clearTimeout(timer);
            errorfn(r, ignore, 9);
        };

        var timeoutFn9 = function(ignore, ignore2, timer) 
        {
            clearTimeout(timer);
            linkToError("ajax timeout", "Timed out with no response.")(document.querySelectorAll("#resultUpdated" + 9)[0].firstChild);
        }

        var successFn10 = function(ignore, r, timer) 
        {
            clearTimeout(timer);
            getUpdateResponse(ignore, r, 10);
        };

        var errorFn10 = function(r, ignore, timer) 
		{
            clearTimeout(timer);
            errorfn(r, ignore, 10);
        };

        var timeoutFn10 = function(ignore, ignore2, timer) 
        {
            clearTimeout(timer);
            linkToError("ajax timeout", "Timed out with no response.")(document.querySelectorAll("#resultUpdated" + 10)[0].firstChild);
        }

        var successFn11 = function(ignore, r, timer) 
        {
            clearTimeout(timer);
            getUpdateResponse(ignore, r, 11);
        };

        var errorFn11 = function(r, ignore, timer) 
		{
            clearTimeout(timer);
            errorfn(r, ignore, 11);
        };

        var timeoutFn11 = function(ignore, ignore2, timer) 
        {
            clearTimeout(timer);
            linkToError("ajax timeout", "Timed out with no response.")(document.querySelectorAll("#resultUpdated" + 11)[0].firstChild);
        }

        var successFn12 = function(ignore, r, timer) 
        {
            clearTimeout(timer);
            getUpdateResponse(ignore, r, 12);
        };

        var errorFn12 = function(r, ignore, timer) 
		{
            clearTimeout(timer);
            errorfn(r, ignore, 12);
        };

        var timeoutFn12 = function(ignore, ignore2, timer) 
        {
            clearTimeout(timer);
            linkToError("ajax timeout", "Timed out with no response.")(document.querySelectorAll("#resultUpdated" + 12)[0].firstChild);
        }

        var successFn13 = function(ignore, r, timer) 
        {
            clearTimeout(timer);
            getUpdateResponse(ignore, r, 13);
        };

        var errorFn13 = function(r, ignore, timer) 
		{
            clearTimeout(timer);
            errorfn(r, ignore, 13);
        };

        var timeoutFn13 = function(ignore, ignore2, timer) 
        {
            clearTimeout(timer);
            linkToError("ajax timeout", "Timed out with no response.")(document.querySelectorAll("#resultUpdated" + 13)[0].firstChild);
        }

        var successFn14 = function(ignore, r, timer) 
        {
            clearTimeout(timer);
            getUpdateResponse(ignore, r, 14);
        };

        var errorFn14 = function(r, ignore, timer) 
		{
            clearTimeout(timer);
            errorfn(r, ignore, 14);
        };

        var timeoutFn14 = function(ignore, ignore2, timer) 
        {
            clearTimeout(timer);
            linkToError("ajax timeout", "Timed out with no response.")(document.querySelectorAll("#resultUpdated" + 14)[0].firstChild);
        }

        var successFn15 = function(ignore, r, timer) 
        {
            clearTimeout(timer);
            getUpdateResponse(ignore, r, 15);
        };

        var errorFn15 = function(r, ignore, timer) 
		{
            clearTimeout(timer);
            errorfn(r, ignore, 15);
        };

        var timeoutFn15 = function(ignore, ignore2, timer) 
        {
            clearTimeout(timer);
            linkToError("ajax timeout", "Timed out with no response.")(document.querySelectorAll("#resultUpdated" + 15)[0].firstChild);
        }

        http[1] = xhr(getURL(1), successFn1, errorFn1, 15000, timeoutFn1);
        errorFn1.timer = http.timer;
        successFn1 = http.timer;
        http[1].get();

        http[2] = xhr(getURL(2), successFn2, errorFn2, 15000, timeoutFn2);
        errorFn2.timer = http.timer;
        successFn2 = http.timer;
        http[2].get();

        http[3] = xhr(getURL(3), successFn3, errorFn3, 15000, timeoutFn3);
        errorFn3.timer = http.timer;
        successFn3 = http.timer;
        http[3].get();

        http[4] = xhr(getURL(4), successFn4, errorFn4, 15000, timeoutFn4);
        errorFn4.timer = http.timer;
        successFn4 = http.timer;
        http[4].get();

        http[5] = xhr(getURL(5), successFn5, errorFn5, 15000, timeoutFn5);
        errorFn5.timer = http.timer;
        successFn5 = http.timer;
        http[5].get();

        http[6] = xhr(getURL(6), successFn6, errorFn6, 15000, timeoutFn6);
        errorFn6.timer = http.timer;
        successFn6 = http.timer;
        http[6].get();

        http[6] = xhr(getURL(6), successFn6, errorFn6, 15000, timeoutFn6);
        errorFn6.timer = http.timer;
        successFn6 = http.timer;
        http[6].get();

        http[7] = xhr(getURL(7), successFn7, errorFn7, 15000, timeoutFn7);
        errorFn7.timer = http.timer;
        successFn7 = http.timer;
        http[7].get();

        http[8] = xhr(getURL(8), successFn8, errorFn8, 15000, timeoutFn8);
        errorFn8.timer = http.timer;
        successFn8 = http.timer;
        http[8].get();

        http[9] = xhr(getURL(9), successFn9, errorFn9, 15000, timeoutFn9);
        errorFn9.timer = http.timer;
        successFn9 = http.timer;
        http[9].get();

        http[10] = xhr(getURL(10), successFn10, errorFn10, 15000, timeoutFn10);
        errorFn10.timer = http.timer;
        successFn10 = http.timer;
        http[10].get();

        http[11] = xhr(getURL(11), successFn11, errorFn11, 15000, timeoutFn11);
        errorFn11.timer = http.timer;
        successFn11 = http.timer;
        http[11].get();

		http[12] = xhr(getURL(12), successFn12, errorFn12, 15000, timeoutFn12);
        errorFn12.timer = http.timer;
        successFn12 = http.timer;
        http[12].get();

        http[13] = xhr(getURL(13), successFn13, errorFn13, 15000, timeoutFn13);
        errorFn13.timer = http.timer;
        successFn13 = http.timer;
        http[13].get();

        http[14] = xhr(getURL(14), successFn14, errorFn14, 15000, timeoutFn14);
        errorFn14.timer = http.timer;
        successFn14 = http.timer;
        http[14].get();

        http[15] = xhr(getURL(15), successFn15, errorFn15, 15000, timeoutFn15);
        errorFn15.timer = http.timer;
        successFn15 = http.timer;
        http[15].get();        
	}
    // function getRowCount() {
    //     var count = 0;
    //     for(var id in countColumn){
    //         count++;
    //     }
    //     return count;
    // }
    // function showAllCount()
    // {
    //     if (displayedContent) return;
    //     for(var id in countColumn){
    //         var value = countColumn[id];
    //         var eleID = id.toString();
    //         if (countState[id] == 1)
    //             document.getElementById(eleID).innerHTML = value;
    //         else if (countState[id] == 2)
    //         {
    //             var element = document.querySelectorAll("#" + id)[0];
    //             var contents = countColumn[id].split(":::");
    //             element.innerHTML = "<span class=\"error\">" + contents[0] + "</span>";
    //             element.style.cursor = "pointer";
    //             element.childNodes[0].style.cursor = "pointer";
    //             element.addEventListener("click", function(e){alert(contents[0] + ":\n\n" + contents[1]);}, false);
    //         }
    //         else if (countState[id] == 3){
    //             var element = document.querySelectorAll("#" + id)[0];
    //             element.innerHTML = countColumn[id];
    //         }
    //     }
    //     displayedContent = true;
    // }
    function getUpdateResponse(ignore, r, id) 
	{
  //       var eleID = 'resultServer'+id.toString();
		// document.getElementById(eleID).innerHTML = r.getElementsByTagName('response_server_name')[0].firstChild.nodeValue;
		// eleID = 'resultDatabase'+id.toString();
		// document.getElementById(eleID).innerHTML = r.getElementsByTagName('response_database_name')[0].firstChild.nodeValue;
        document.getElementById('resultUpdated' + id).innerHTML = r.getElementsByTagName('response_records_updated')[0].firstChild.nodeValue;
        // countState['resultUpdated' + id] = 1;
        // countColumn['resultUpdated' + id] = r.getElementsByTagName('response_records_updated')[0].firstChild.nodeValue;
        // if (getRowCount() == locationCount)
        // {
        //     showAllCount();
        // }
	}
	function linkToError(title, message) {
		return function showError(e) {
	        var parent = e.parentNode;
            var id = parent.getAttribute('id');
            var element = document.querySelectorAll("#" + id)[0];
            element.innerHTML = "<span class=\"error\">" + title + "</span>";
            element.style.cursor = "pointer";
            element.childNodes[0].style.cursor = "pointer";
            element.addEventListener("click", function(e){alert(title + ":\n\n" + message);}, false);
            // countState[id] = 2;
            // countColumn[id] = title + ":::" + message;
            // if (getRowCount() == locationCount)
            // {
            //     showAllCount();
            // }
	   	}
   	}
   	function errorfn(r, ignore, ServerID) 
    {
	    var element = document.querySelectorAll("#resultUpdated" + ServerID)[0];
	    if (r && r.length > 0) {
            linkToError("ajax error", r.toString())(element.firstChild);}
	    else if (!element.marked) {
            // countState["resultUpdated" + ServerID] = 3;
            // countColumn["resultUpdated" + ServerID] = "<span class=\"error\">ajax error: no details.</span>";
            // if (getRowCount() == locationCount)
            // {
            //     showAllCount();
            // }
            element.innerHTML = "<span class=\"error\">ajax error: no details.</span>";
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
		// var element = document.createElement("textarea");
		// document.body.appendChild(element);
		// element.style.width = "100%";
		// element.innerHTML = body;
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
			<td class="subHead" width="30%">
			  Server
			</td>
			<td class="subHead" width="30%">
			  Database
			</td>
			<td class="subHead" width="25%">
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
