<!DOCTYPE html>

<!--#include virtual="ha_backoffice/includes/Config.asp"-->
<!--#include virtual="ha_backoffice/Servers/TableDistribution/includes/TableDistribution_ReadObjects.inc"-->

<%
	TableName = "z2t_Activity"
	Call ReadObjects(TableName)
%>

<html>
<head>
    <title>Table Distribution Status - <%=TableName%></title>
	
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>ha_BackofficePopup.css">

    <script language="javascript" type="text/javascript" src="<%=strPathBOIncludes%>haFunctions.js"></script>
    <script language="javascript" type="text/javascript" src="<%=strPath%>/includes/UserTracking/UserTrackingPost.js"></script>
	
<script language='JavaScript'>

	var http = new XMLHttpRequest();
	var http2 = new XMLHttpRequest();
	var loopStep = 1;
	var locationCount = <%=LocationCount%>;
	var locationServer = ['',  
		<%For i = 0 to LocationCount
			Response.write "'" & LocationServer(i) & "', "
		Next%>];
	var locationDatabase = ['', 
		<%For i = 0 to LocationCount
			Response.write "'" & LocationDatabase(i) & "', "
		Next%>];
	var locationTable = ['',  
		<%For i = 0 to LocationCount
			Response.write "'" & LocationTable(i) & "', "
		Next%>];

	function addCommas(nStr)
		{
		nStr += '';
		x = nStr.split('.');
		x1 = x[0];
		x2 = x.length > 1 ? '.' + x[1] : '';
		var rgx = /(\d+)(\d{3})/;
		while (rgx.test(x1)) 
			{
			x1 = x1.replace(rgx, '$1' + ',' + '$2');
			}
		return x1 + x2;
		}
		
    function clearResults()
        {
		for (i=1; i < locationCount + 2; i++)
			{
			var eleID = 'resultServer'+i.toString();
			document.getElementById(eleID).innerHTML = '- - - - -';
			var eleID = 'resultDatabase'+i.toString();
			document.getElementById(eleID).innerHTML = '- - - - -';
			var eleID = 'resultTable'+i.toString();
			document.getElementById(eleID).innerHTML = '- - - - -';
			var eleID = 'resultRecords'+i.toString();
			document.getElementById(eleID).innerHTML = '- - - - -';
			var eleID = 'resultLocation1'+i.toString();
			document.getElementById(eleID).innerHTML = '- - - - -';
			var eleID = 'resultLocation2'+i.toString();
			document.getElementById(eleID).innerHTML = '- - - - -';
			}
		document.getElementById('Destination1CountCollectionTable').innerHTML = '- - - - -';
		document.getElementById('Destination1CountAnnualTable').innerHTML = '- - - - -';
		document.getElementById('Destination2CountCollectionTable').innerHTML = '- - - - -';
		document.getElementById('Destination2CountAnnualTable').innerHTML = '- - - - -';
		}
		
    function clickRun()
        {
		clearResults();
		loopStep = 1;
        runStatusCheck();
		runDestinationCount();
		}
		
    function formLoad()
        {
        SetScreen(1100, 750);
		UserTracking('formLoad','TableDistribution_<%=TableName%>.asp','','');
		}
		
    function runStatusCheck() 
        {
		var url = '<%=strPathServers%>TableDistribution/z2t_Activity/TableDistribution_z2t_Activity_readstatus.asp' + 
			'?s=' + locationServer[loopStep] +
			'&d=' + locationDatabase[loopStep] +
			'&t=' + locationTable[loopStep] +
			'&Now=' + escape(Date())
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
				var eleID = 'resultTable'+loopStep.toString();
				document.getElementById(eleID).innerHTML = r.getElementsByTagName('response_table_name')[0].firstChild.nodeValue;
				var eleID = 'resultRecords'+loopStep.toString();
				document.getElementById(eleID).innerHTML = addCommas(r.getElementsByTagName('response_record_count')[0].firstChild.nodeValue);
				var eleID = 'resultLocation1'+loopStep.toString();
				document.getElementById(eleID).innerHTML = addCommas(r.getElementsByTagName('response_location_1')[0].firstChild.nodeValue);
				var eleID = 'resultLocation2'+loopStep.toString();
				document.getElementById(eleID).innerHTML = addCommas(r.getElementsByTagName('response_location_2')[0].firstChild.nodeValue);
				
				if (loopStep < locationCount + 1)
					{
					loopStep = loopStep + 1;
					runStatusCheck();
					}
				}
			}
        }

    function runDestinationCount() 
        {
		var url = '<%=strPathServers%>TableDistribution/z2t_Activity/TableDistribution_z2t_Activity_Destinations_readstatus.asp' + 
			'?Now=' + escape(Date())
		//alert(url);
			
        http2.open('GET', url, true);
        http2.onreadystatechange = getDestinationResponse;
        http2.send();
        }

    function getDestinationResponse() 
        {
        if (http2.readyState == 4) 
			{
			if (http2.status == 200) 
				{
				var r = http2.responseXML;
				var rd = r.getElementsByTagName("activity_destinations")[0].childNodes
				
				for (i=0; i < rd.length; i++)
					{
					var eleID = 'Destination' + (i+1) + 'CountCollectionTable';
						document.getElementById(eleID).innerHTML = addCommas(rd[i].childNodes[2].childNodes[0].nodeValue);
					var eleID = 'Destination' + (i+1) + 'CountAnnualTable';
						document.getElementById(eleID).innerHTML = addCommas(rd[i].childNodes[3].childNodes[0].nodeValue);
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

	<!--#include virtual="ha_backoffice/Servers/TableDistribution/includes/TableDistribution_Page_heading.inc"-->
		  
		<table width="100%" border="0" cellspacing="2" cellpadding="2">
		
          <tr>
		    <td class="databaseHead" colspan="3">
			  Locations
			</td>
			<td>
			  &nbsp;
			</td>
		    <td class="databaseHead" colspan="6">
			  Results
			</td>
		  </tr>
		  
          <tr>
		    <td class="subHead2" colspan="8">
			  &nbsp;
			</td>
		    <td class="subHead2">
			  Transferred
			</td>
		    <td class="subHead2">
			  Transferred
			</td>
		  </tr>
		  
          <tr>
            <td width="3%">
			  &nbsp;
			</td>
            <td class="subHead" width="8%">
              Server
            </td>
            <td class="subHead" width="14%">
              Database
            </td>
            <td width="2%">
			  &nbsp;
			</td>
            <td class="subHead" width="8%">
			  Server
			</td>
            <td class="subHead" width="14%">
			  Database
			</td>
            <td class="subHead" width="20%">
			  Table
			</td>
            <td class="subHead" width="9%">
			  Records
			</td>
            <td class="subHead" width="9%">
			  Location 1
			</td>
            <td class="subHead" width="9%">
			  Location 2
			</td>
          </tr>
		  
		  
<%
	For i = 1 to LocationCount + 1
%>
		  
          <tr>
            <td align="right">
			  <span style="color: #C0C0C0;font-size: 10px;"><%=cStr(i)%>.</span>
            </td>
            <td>
              <%=LocationServer(i-1)%>
            </td>
            <td>
              <%=LocationDatabase(i-1)%>
            </td>
			<td>
			  &nbsp;
			</td>
			<td id="resultServer<%=i%>">
			  - - - - -
			</td>
			<td id="resultDatabase<%=i%>">
			  - - - - -
			</td>
			<td id="resultTable<%=i%>">
			  - - - - -
			</td>
			<td id="resultRecords<%=i%>" align="right">
			  - - - - -
			</td>
			<td id="resultLocation1<%=i%>" align="right">
			  - - - - -
			</td>
			<td id="resultLocation2<%=i%>" align="right">
			  - - - - -
			</td>
          </tr>
		  
<%
	Next
%>
          <tr>
		    <td colspan="10">
			  <hr>
			</td>
		  </tr>

        </table>	
		
		<table width="100%" border="0" cellspacing="2" cellpadding="2">
          <tr>
		    <td class="databaseHead" colspan="3">
			  Destinations
			</td>
			<td>
			  &nbsp;
			</td>
		    <td class="databaseHead" colspan="2">
			  Counts
			</td>
		  </tr>

          <tr>
            <td colspan="4">
			  &nbsp;
			</td>
            <td class="subHead2">
			  Collection
			</td>
            <td class="subHead2">
			  Annual
			</td>
          </tr>
		  
          <tr>
            <td class="subHead" width="20%">
			  Description
			</td>
            <td class="subHead" width="25%">
              Server
            </td>
            <td class="subHead" width="25%">
              Database
            </td>
            <td width="4%">
			  &nbsp;
			</td>
            <td class="subHead" width="13%">
			  Table
			</td>
            <td class="subHead" width="13%">
			  Table
			</td>
          </tr>

		  <tr>
		    <td>
			  Transfer Destination 1
			</td>
		    <td>
			  Philly05
			</td>
		    <td>
			  z2t_Backoffice
			</td>
            <td>
			  &nbsp;
			</td>
			<td id="Destination1CountCollectionTable" align="right">
			  - - - - -
			</td>
			<td id="Destination1CountAnnualTable" align="right">
			  - - - - -
			</td>
		  </tr>
		  <tr>
		    <td>
			  Transfer Destination 2
			</td>
		    <td>
			  Casper10
			</td>
		    <td>
			  z2t_Backoffice
			</td>
            <td>
			  &nbsp;
			</td>
			<td id="Destination2CountCollectionTable" align="right">
			  - - - - -
			</td>
			<td id="Destination2CountAnnualTable" align="right">
			  - - - - -
			</td>
		  </tr>
		  
          <tr>
		    <td colspan="6">
			  <hr>
			</td>
		  </tr>
		</table>
  
        <table width="100%" border="0" cellspacing="5" cellpadding="5">
		  <tr>
		    <td width="45%">
			  &nbsp;
			</td>
		    <td align="center">
                <a href="javascript:clearResults();" class="bo_Button100" title="Clear the Results">Clear</a>
		    </td>
		    <td align="center">
                <a href="javascript:clickRun();" class="bo_Button100" title="Lists the status of the tables">Run Status</a>
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

</body>
</html>
