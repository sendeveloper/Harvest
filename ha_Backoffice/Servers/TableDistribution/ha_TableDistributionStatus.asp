<!DOCTYPE html>

<!--#include virtual="ha_backoffice/includes/Config.asp"-->
<!--#include virtual="ha_backoffice/Servers/TableDistribution/includes/TableDistribution_ReadObjects.inc"-->

<%			
	TableName = RequestParse(Request("t"),"")
	Increment = RequestParse(Request("i"),10000)
	
	UpdateClass = RequestParse(Request("u"),"")
	Select Case UpdateClass
	Case "d"
		UpdateClass = " (Development)"
	Case "s"
		UpdateClass = " (Staging)"
	Case "o"
		UpdateClass = " (Promotion)"
	Case "p"
		UpdateClass = " (Production)"
	Case "a"
		UpdateClass = " (Archive)"
	Case Else
		UpdateClass = "" 
	End Select

	ObjectsClass = RequestParse(Request("o"),"")
	Call ReadObjects(ObjectsClass)
	'for z2t_SpecialDistricts we need to change the table name
	If left(TableName, 19) = "z2t_SpecialDistrict" Then
		For i = 0 to LocationCount
			LocationTable(i) = TableName
		Next 
	End If
	
	RequestStyle = "Generic"
	
	Dim StartEndData
	StartEndData = "&Start=0&End=100"
%>

<html>
<head>
    <title>Table Distribution Status - <%=TableName%><%=UpdateClass%></title>

    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>ha_BackofficePopup.css">

    <script language="javascript" type="text/javascript" src="<%=strPathBOIncludes%>haFunctions.js"></script>
    <script language="javascript" type="text/javascript" src="<%=strPath%>/includes/UserTracking/UserTrackingPost.js"></script>
		
<script language='JavaScript'>

	var http = new XMLHttpRequest();
	var tableName = '<%=TableName%>'
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
	var locationTable = [ 
		<%For i = 0 to LocationCount
			Response.write "'" & LocationTable(i) & "', "
		Next%>];
		
	var increment = parseInt(<%=Increment%>);
	var aRecStart = 0;
	var aRecEnd = 0;
	var aRecMax = 0;
	var aRecMin = 0;
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
		var url = '<%=strPathServers%>TableDistribution/ha_TableDistribution_readstatus.asp' +
			'?s=' + locationServer[loopStep] +
			'&d=' + locationDatabase[loopStep] +
			'&t=' + locationTable[loopStep] +
			'&ExecBy=TableDistributionStatus_<%=TableName%>.asp' +
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
					aRecMin = parseInt(r.getElementsByTagName('response_minimum_ID')[0].firstChild.nodeValue);
					aRecMax = parseInt(r.getElementsByTagName('response_maximum_ID')[0].firstChild.nodeValue);
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
					aRecStart = aRecMin;
					aRecEnd = aRecMin + increment - 1;
					totalInserted = 0;
					totalUpdated = 0;
					runUpdate();
					}				
				}
			}
        }
		
	function runUpdate() 
        {		
		var url = '<%=strPathServers%>TableDistribution/ha_TableDistribution_update.asp' +
			'?s=' + locationServer[loopStep] +
			'&d=' + locationDatabase[loopStep] +
			'&t=' + locationTable[loopStep] +
			'&o=<%=ObjectsClass%>' +
			'&Start=' + aRecStart +
			'&End=' + aRecEnd +
			'&ExecBy=TableDistributionStatus_<%=TableName%>.asp' +
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
				
				if (aRecEnd < aRecMax)
					{
					//Send it back for more records
					aRecStart = aRecEnd + 1;
					aRecEnd = aRecEnd + increment;
					runUpdate();
					}
				else
					{
					if (loopStep < locationCount)
						{
						//Table finished, move to the next
						loopStep = loopStep + 1;
						aRecStart = aRecMin;
						aRecEnd = aRecMin + increment - 1;
						totalInserted = 0;
						totalUpdated = 0;
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
		
    </td>
  </tr>
</table>

</body>
</html>
