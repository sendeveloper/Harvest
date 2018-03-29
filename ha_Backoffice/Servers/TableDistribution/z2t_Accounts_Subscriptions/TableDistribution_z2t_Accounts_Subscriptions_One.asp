<!DOCTYPE html>

<!--#include virtual="ha_backoffice/includes/Config.asp"-->
<!--#include virtual="ha_backoffice/Servers/TableDistribution/includes/TableDistribution_ReadObjects.inc"-->

<%
	If isnull(Request("z2tID")) or Request("z2tID") = "" Then
		z2tID = ""
	Else
		z2tID = Request("z2tID")
	End If
	
	TableName = "z2t_Accounts_Subscriptions"
	Call ReadObjects(TableName)
%>

<html>
<head>
    <title>Table Distribution Status z2t_Accounts_Subscriptions</title>

    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>ha_BackofficePopup.css">

    <script language="javascript" type="text/javascript" src="<%=strPathBOIncludes%>haFunctions.js"></script>
		
<script language='JavaScript'>

	var http = new XMLHttpRequest();
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
				
	function runUpdate() 
        {		
		var url = '<%=strPathServers%>TableDistribution/z2t_Accounts_Subscriptions/TableDistribution_z2t_Accounts_Subscriptions_update.asp' +
			'?s=' + locationServer[loopStep] +
			'&d=' + locationDatabase[loopStep] +
			'&z2tID=' + z2tID +
			'&ExecBy=TableDistribution_z2t_Accounts_Subscriptions_One.asp' + 
			'&Now=' + escape(Date());
		//alert(url);
		var eleID = 'resultServer'+loopStep;
        document.getElementById(eleID).innerHTML = locationServer[loopStep];
        eleID = 'resultDatabase'+loopStep;
        document.getElementById(eleID).innerHTML = locationDatabase[loopStep];
			
        http.open('GET', url, true);
        http.onreadystatechange = getUpdateResponse;
        http.send();
        }

    function getUpdateResponse() 
        {
        if (http.readyState == 4) 
			{
			if (http.status == 200) 
				{
				var r = http.responseXML;
				// var eleID = 'resultServer'+loopStep.toString();
				// document.getElementById(eleID).innerHTML = r.getElementsByTagName('response_server_name')[0].firstChild.nodeValue;
				// var eleID = 'resultDatabase'+loopStep.toString();
				// document.getElementById(eleID).innerHTML = r.getElementsByTagName('response_database_name')[0].firstChild.nodeValue;
				var eleID = 'resultUpdated'+loopStep.toString();
				document.getElementById(eleID).innerHTML = r.getElementsByTagName('response_records_updated')[0].firstChild.nodeValue;
				
				if (loopStep < locationCount)
					{
					//Table finished, move to the next
					loopStep = loopStep + 1;
					runUpdate();
					}
				else
					{
					//All done
					//window.close()
					
					//Now we run the Addons update
					var url = 'http://www.harvestamerican.info/ha_backoffice/Servers/TableDistribution/z2t_Accounts_Subscriptions_Addons/' +
						'TableDistribution_z2t_Accounts_Subscriptions_Addons_One.asp?z2tID=' + z2tID;
					location.href = url;
					}
				}
			}
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
              Update - z2t_Accounts_Subscriptions
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
