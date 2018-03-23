<!--#include virtual="includes/connection.asp"-->
<!--#include virtual="ha_backoffice/includes/Config.asp"-->
<%
	Dim rs
    Dim SQL
	'Dim connDallas
	Dim DomainName
		
	RequestError = 0

	If isnull(Request("d")) or Request("d") = "" Then
		RequestError = 1
	Else
		DomainName = Request("d")
	End If
    
%>

<html>
<head>
    <title>Domain Status - <%=DomainName%></title>

    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <link href="/ha_BackOffice/includes/ha_BackofficePopup.css" type="text/css" rel="stylesheet" />
    <script src="/ha_BackOffice/includes/haFunctions.js"  type="text/javascript" language="javascript"></script>

	<script type="text/javaScript">

		var http = new XMLHttpRequest();

		function formLoad()
			{
			SetScreen(850, 550);
			makeCorsRequest();
			}

		// Create the XHR object.
		function createCORSRequest(method, url) 
			{
			var xhr = new XMLHttpRequest();
			if ("withCredentials" in xhr) 
				{
				// XHR for Chrome/Firefox/Opera/Safari.
				xhr.open(method, url, true);
				} 
			else if (typeof XDomainRequest != "undefined") 
				{
				// XDomainRequest for IE.
				xhr = new XDomainRequest();
				xhr.open(method, url);
				} 
			else 
				{
				// CORS not supported.
				xhr = null;
				}
			return xhr;
			}

		// Make the actual CORS request.
		function makeCorsRequest() 
			{
			var xhr = createCORSRequest('GET', '<%=DomainName%>/ha_DomainTest.asp');
			if (!xhr) 
				{
				alert('CORS not supported');
				return;
				}

			// Response handlers.
			xhr.onload = function() 
				{
				var text = xhr.responseText;
				
				var r = xhr.responseXML;
				
				var p = r.getElementsByTagName('Domain_Test')[0].childNodes
				
				for (var loopStep = 0; loopStep < p.length; loopStep++) 
					{
					var eleID = 'resultItem'+(loopStep+1).toString();
					document.getElementById(eleID).innerHTML = p[loopStep].nodeName;
					var eleID = 'resultResponse'+(loopStep+1).toString();
					document.getElementById(eleID).innerHTML = p[loopStep].firstChild.nodeValue;
					}
				};

			xhr.onerror = function() 
				{
				alert('Woops, there was an error making the request.');
				};

			xhr.send();
			}
		
	</script>

</head>

<body onload="formLoad();">

<table width="800" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td>

        <table width="100%" border="0" cellspacing="10" cellpadding="0" align="center">

          <tr>
            <td class="popupHeading">
              Domain Status - <%=DomainName%>
            </td>
          </tr>

		  <tr>
		    <td>
		  
			  <table width="100%" border="0" cellspacing="2" cellpadding="2">
				
				  <tr>
					<td width="35%" class="popupHeading2">
					  Item
					</td>
					<td width="5%">
					  &nbsp;
					</td>
					<td width="60%" class="popupHeading2">
					  Response
					</td>
				  </tr>
				  
<%
	For i = 1 to 10
%>
				<tr>
				  <td id="resultItem<%=i%>" align="right">
					- - - - -
				  </td>
				  <td>
				    &nbsp;
				  </td>
				  <td id="resultResponse<%=i%>">
					- - - - -
				  </td>
				</tr>
				
<%
	Next
%>
				  <tr>
					<td class="popupHeading2">
					  &nbsp;
					</td>
					<td>
					  &nbsp;
					</td>
					<td class="popupHeading2">
					  &nbsp;
					</td>
				  </tr>
				
			  </table>	
				
		 
				<table width="100%" border="0" cellspacing="5" cellpadding="5">
				  <tr>
					<td width="45%">
					  &nbsp;
					</td>

					<td align="center">
						<a href="javascript:close();" class="bo_Button100" title="Closes this window">Close</a>
					</td>
					<td width="45%">
					  &nbsp;
					</td>
				  </tr>
				</table>

            </td>
          </tr>
        </table>	
		
    </td>
  </tr>
</table>
</body>
</html>

<script type="text/javascript">
    window.opener.location.href = window.opener.location.href;
</script>

<%
	If left(DomainName,8) = "https://" Then
		DN = mid(DomainName,9)
	Else
		DN = mid(DomainName,8)
	End If

	SQL = "INSERT INTO ha_BackOffice.dbo.ha_TaskLog " & _
        "(CreatedBy, CreatedDate, Task, Status) " & _
		"VALUES " & _
        "('ha_Domain_readstatus.asp', GETDATE(), '" & DN & "', 'Finished')"
		
	connCasper10.Execute SQL
%>

