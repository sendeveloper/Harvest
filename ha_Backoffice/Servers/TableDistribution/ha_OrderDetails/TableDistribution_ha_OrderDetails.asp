<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

<!--#include virtual="ha_backoffice/includes/Config.asp"-->
<!--#include virtual="ha_backoffice/Servers/TableDistribution/includes/TableDistribution_ReadObjects.inc"-->

<%
	TableName = "ha_OrderDetails"
	Call ReadObjects(TableName)%>

<html>
<head>
    <title>Table Distribution Status - <%=TableName%></title>

    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>ha_BackofficePopup.css">

    <script language="javascript" type="text/javascript" src="<%=strPathBOIncludes%>haFunctions.js"></script>
    <script language="javascript" type="text/javascript" src="<%=strPath%>/includes/UserTracking/UserTrackingPost.js"></script>
		
<script language='JavaScript'>

	var http = new XMLHttpRequest();
	var loopStep = 0;
	var locationCount = <%=LocationCount%>;
	var locationServer = [ 
		<%For i = 0 to LocationCount
			Response.write "'" & LocationServer(i) & "', "
		Next%>];
	var locationDatabase = [ 
		<%For i = 0 to LocationCount
			Response.write "'" & LocationDatabase(i) & "', "
		Next%>];
	var aRecSkip = 140000;
	var aRecStart = 0;
	var aRecEnd = 0;
	var aRecMax = 0;
	var updateFlag = 0;
	var totalInserted = 0;
	var totalUpdated = 0;

    function clearResults()
        {
		loopStep = 0;
		for (i=0; i < locationCount + 1; i++)
			{
			var eleID = 'resultServer'+i.toString();
			document.getElementById(eleID).innerHTML = '- - - - -';
			var eleID = 'resultDatabase'+i.toString();
			document.getElementById(eleID).innerHTML = '- - - - -';
			var eleID = 'resultRecords'+i.toString();
			document.getElementById(eleID).innerHTML = '- - - - -';
			
			if (i > 0)
				{
				var eleID = 'resultInserted'+i.toString();
				document.getElementById(eleID).innerHTML = '- - - - -';
				var eleID = 'resultUpdated'+i.toString();
				document.getElementById(eleID).innerHTML = '- - - - -';
				}
			}
		}
		
    function clickRun()
        {
		clearResults();
        runStatusCheck();
		}
		
    function clickUpdate()
        {
		clearResults();
        runStatusCheck();
		updateFlag = 1;
		}
		
    function formLoad()
        {
        SetScreen(1100, 600);
		UserTracking('formLoad','TableDistribution_<%=TableName%>.asp','','');		
		}
				
    function runStatusCheck() 
        {		
		var url = '<%=strPathServers%>TableDistribution/' +
			'ha_OrderDetails/TableDistribution_ha_OrderDetails_readstatus.asp' +
			'?s=' + locationServer[loopStep] +
			'&d=' + locationDatabase[loopStep] +
			'&Now=' + escape(Date());
		//alert(url);
			
        http.open('GET', url, true);
        http.onreadystatechange = getResponse;
        http.send();
        }

    function getResponse() 
        {
        if (http.readyState == 4) 
			{
			if (http.status == 200) 
				{
				var r = http.responseXML;
				var eleID = 'resultServer'+loopStep.toString();
				document.getElementById(eleID).innerHTML = r.getElementsByTagName('response_server_name')[0].firstChild.nodeValue;
				var eleID = 'resultDatabase'+loopStep.toString();
				document.getElementById(eleID).innerHTML = r.getElementsByTagName('response_database_name')[0].firstChild.nodeValue;
				var eleID = 'resultRecords'+loopStep.toString();
				document.getElementById(eleID).innerHTML = r.getElementsByTagName('response_record_count')[0].firstChild.nodeValue;

				//Set the maximum AccountID from the source response
				if (loopStep == 0)
					{
					aRecMax = r.getElementsByTagName('response_maximum_ODID')[0].firstChild.nodeValue;
					//alert(aRecMax);
					}
					
				//Move to the next location
				if (loopStep < locationCount)
					{
					loopStep = loopStep + 1;
					runStatusCheck();
					}
				else if (updateFlag == 1)
					{
					//Start the update portion
					loopStep = 1;  //We don't want to start at 0, that's the source
					aRecStart = aRecSkip + 1;
					aRecEnd = aRecStart + 1999;
					//alert('aRecStart=' + aRecStart);
					//alert('aRecEnd=' + aRecEnd);
					totalInserted = 0;
					totalUpdated = 0;
					runUpdate();
					}				
				}
			}
        }
		
	function runUpdate() 
        {		
		var url = '<%=strPathServers%>TableDistribution/' +
			'ha_OrderDetails/TableDistribution_ha_OrderDetails_update.asp' +
			'?s=' + locationServer[loopStep] +
			'&d=' + locationDatabase[loopStep] +
			'&Start=' + aRecStart +
			'&End=' + aRecEnd +
			'&ExecBy=TableDistribution_ha_OrderDetails.asp' +
			'&Now=' + escape(Date());
		//alert(url);
			
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
				var eleID = 'resultServer'+loopStep.toString();
				document.getElementById(eleID).innerHTML = r.getElementsByTagName('response_server_name')[0].firstChild.nodeValue;
				var eleID = 'resultDatabase'+loopStep.toString();
				document.getElementById(eleID).innerHTML = r.getElementsByTagName('response_database_name')[0].firstChild.nodeValue;
				
				totalInserted += parseInt(r.getElementsByTagName('response_records_inserted')[0].firstChild.nodeValue);
				var eleID = 'resultInserted'+loopStep.toString();
				document.getElementById(eleID).innerHTML = totalInserted;
				
				totalUpdated += parseInt(r.getElementsByTagName('response_records_updated')[0].firstChild.nodeValue);
				var eleID = 'resultUpdated'+loopStep.toString();
				document.getElementById(eleID).innerHTML = totalUpdated;
				
				//alert('aRecEnd=' + aRecEnd);
				//alert('aRecMax=' + aRecMax);
				
				if (aRecEnd < aRecMax)
					{
					//Send it back for more records
					aRecStart = aRecEnd + 1;
					aRecEnd = aRecEnd + 2000;
					runUpdate();
					}
				else
					{
					if (loopStep < locationCount)
						{
						//Table finished, move to the next
						loopStep = loopStep + 1;
						runUpdate();
						}
					else
						{
						//All done
						updateFlag = 0;
						}
					}
				}
			}
        }

</script>

</head>

<body onLoad="formLoad();">

<table width="1000" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td>

	<!--#include virtual="ha_backoffice/Servers/TableDistribution/includes/TableDistribution_Page_top.inc"-->
	<!--#include virtual="ha_backoffice/Servers/TableDistribution/includes/TableDistribution_Page_loop.inc"-->
	<!--#include virtual="ha_backoffice/Servers/TableDistribution/includes/TableDistribution_Page_bottom.inc"-->
	<!--#include virtual="ha_backoffice/Servers/TableDistribution/includes/TableDistribution_Page_buttons.inc"-->

		<table border="0" cellspacing="2" cellpadding="2" align="center">		
          <tr>
		    <td>
			  <br>
			  Note: Only the last 2 years of Order Details are copied
            </td>
          </tr>
        </table>	
		
    </td>
  </tr>
</table>

</body>
</html>
