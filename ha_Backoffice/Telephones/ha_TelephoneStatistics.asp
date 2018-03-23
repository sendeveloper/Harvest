<!doctype html>

<html>
<head>
  <%
	Session("Redirect") = ""
	ColorTab = 2
	PageHeading = "Telephone Statistics"
	
	If Request("submitted") > "" Then
		Submitted = Request("submitted")
	Else
		Submitted = ""
	End If
	
	If Request("extension") > "" Then
		Extension = Request("extension")
	Else
		Extension = ""
	End If

	If Request("extName") > "" Then
		ExtName = Request("extName")
	Else
		ExtName = ""
	End If

	If Request("startDate") > "" Then
		StartDate = Request("startDate")
	Else
		StartDate = date()
	End If

	If Request("endDate") > "" Then
		EndDate = Request("endDate")
	Else
		EndDate = date()
	End If

	If Submitted = "Y" Then
		Fltr = " And CONVERT(datetime, substring(calldate,1,8), 1) Between ''" & StartDate & "'' and ''" & EndDate & "''"

		If CDbl(Extension) = -2 then
			Fltr = Fltr + " And Ring > ''0''''14''"
		ElseIf CDbl(Extension) > -1 then
			Fltr = Fltr + " And LTrim(Ext) = ''" & Extension & "''"
		End If
	Else
		Fltr = ""
	End If
  %>
  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_BackOffice/includes/config.asp" -->
  <title>Harvest American Phone Call Log Statistics</title>
  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />

  <script language="JavaScript" src="<%=strPathControlPanel%>includes/checkDate.js" type="text/javascript"></script>
  <script language="JavaScript" src="datepick/ts_picker.js" type="text/javascript"></script>

  <link href="<%=strPathBOIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet" />
  <link href="<%=strPathControlPanel%>includes/HarvestAmerican.css" type="text/css" rel="stylesheet" />
  <link rel="stylesheet" type="text/css" href="<%=strPathBOIncludes%>menu/dropdowntabfiles/halfmoontabs.css" />
  <script type="text/javascript" src="<%=strPathBOIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>

  <script type="text/javaScript">
    function isDate(sDate) {
      var scratch = new Date(sDate);
      if (scratch.toString() == "NaN" || scratch.toString() == "Invalid Date") {
        alert("Not a Date");
        return false;
      } 
      else 
      {
        return true;
      }
    }
    function changeDates(d)
        {
        document.frm.startDate.value = d;
        document.frm.endDate.value = d;
        clickSubmit();
        }

	function checkStartDate()
		{
		var d = document.frm.startDate;

		var StartDate = d.value;
		var NowDate   = new Date();
		var CurrYear  = NowDate.getFullYear();

		// Fill in the start year if it's missing
		var i=StartDate.indexOf('/');

		if (i > -1)
			{i = StartDate.indexOf('/',i+1);

			if (i > -1)
				{ // Fill in the century if it's missing
				if (StartDate.substring(i+1).length == 2)
					StartDate = StartDate.substring(0,i+1) + '20' + StartDate.substring(i+1)
				}
			else
				StartDate = StartDate + '/' + CurrYear;
			}
		
		if (StartDate.length != 0) 
			{
			if (isDate(StartDate)==false) {
				d.focus();
				return false
				}
			document.getElementById("startDate").value = StartDate;
			return true
			}

		alert('Please enter a start date');
		d.focus();
		return false
		}    

	function checkEndDate()
        {
		var d = document.frm.endDate;

		var EndDate   = d.value;
		var NowDate   = new Date();
		var CurrYear  = NowDate.getFullYear();

		// Fill in the end year if it's missing
		var i=EndDate.indexOf('/');

		if (i > -1)
			{i = EndDate.indexOf('/',i+1);

			if (i > -1)
				{ // Fill in the century if it's missing
				if (EndDate.substring(i+1).length == 2)
					EndDate = EndDate.substring(0,i+1) + '20' + EndDate.substring(i+1)
				}
			else
				EndDate = EndDate + '/' + CurrYear;
			}
		
		if (EndDate.length != 0) 
			{
			if (isDate(EndDate)==false) {
				d.focus();
				return false
				}
			document.getElementById("endDate").value = EndDate;
			return true
			}
			
		alert('Please enter an end date');
		d.focus();
		return false
		}    

	function clickSubmit()
		{
            if (checkStartDate() && checkEndDate()) {
			//Save the extension name (as displayed to the user)
			var obj = document.getElementById("extension");
			document.getElementById("extName").value = obj.options[obj.selectedIndex].text;

            document.frm.submit();
            }
		}

	var isIE = 1;    
	if (navigator.appName == "Netscape")
		{ 
		isIE = 0;
		}

	function whichKey(e)
		{
		if (isIE)
			{
			return e.keyCode;
			}
		else
			{
			return e.which;
			}
		  }
			
	function keytest(e) 
		{
		var k = whichKey(e);
		if (k == 13) 
			{
			clickSubmit();
			}
		}
  </script>

  <!--Load the AJAX API for Google's Chart-->
  <script type="text/javascript" src="https://www.google.com/jsapi"></script>
  <% If Submitted = "Y" Then %>
    <script type="text/javascript">
	  var data;
	  var SeriesNames = new Array();
	  var ChartColors = new Array();
	  
      // Load the Visualization API and the column chart package.
      google.load('visualization', '1.0', {'packages':['corechart']});
      
      // Set a callback to run when the Google Visualization API is loaded.
      google.setOnLoadCallback(drawCharts);

	  // Keep track of the chart series (e.g. phone extensions)
	  function SeriesNbr(nm, clr) {
		var i;
		
		for (i = 0; i < SeriesNames.length; i++)
			{if (SeriesNames[i] == nm)
				return i + 1;
			}

		// Add series to chart
		data.addColumn('number', nm);

		// Add series name to our own array
		i = SeriesNames.length + 1;
		SeriesNames.length = i;
		SeriesNames[i-1] = nm;

		// Add chart color to our array
		ChartColors.length = i;
		ChartColors[i-1] = clr;
		
		return i;
	  }

      // Callback that creates and populates a data table, 
      // instantiates the column charts, passes in the data and draws it.
      function drawCharts() {
		drawChart_01();
		drawChart_02();
	  }
	  
      function drawChart_01() {
		var ChartX = -1;
		SeriesNames = new Array();
		ChartColors = new Array();

		// Create the data table.
		data = new google.visualization.DataTable();
		data.addColumn('string', 'CallHour');

	  <%
		set rs=Server.CreateObject("ADODB.Recordset")
		rs.CursorLocation = 3

		strSQL = "ha_Phone_CallLog_chart_01 ('" & Fltr & "')"
        
		rs.open strSQL, connCasper10, 3, 3, 4
		
		PrevHour = ""
		RecCnt = 0
		
		Do until rs.eof
			CallHour = rs("CallHour")

			If CallHour <> PrevHour Then %>
				ChartX = ChartX + 1;
				data.addRows(1);
				data.setValue(ChartX, 0, '<%=CallHour%>');
		<%		PrevHour = CallHour
			End If
		
			If rs("PhoneLine") > "" Then %>
				data.setValue(ChartX, SeriesNbr("<%=rs("PhoneLine")%>", "<%=rs("Color")%>"), <%=rs("CallCount")%>);
		<%		RecCnt = RecCnt + 1
			End If
			
			rs.MoveNext
		Loop

		rs.Close
		set rs = nothing
		
		If RecCnt > 0 Then %>
		  // Set chart options
		  var options = {'title':'<%=ExtName%>',
						 'isStacked':true,
						 'hAxis':{showTextEvery:1, slantedText:false, maxAlternation:2, textStyle:{fontSize:10}},
						 'chartArea':{left:50, top:25, width:670, height:250},
						 'colors':ChartColors,
						 'backgroundColor':{stroke:"black", strokeWidth:1},
						 'width':870,
						 'height':340
						};

		  // Instantiate and draw our chart, passing in some options.
		  var chart = new google.visualization.ColumnChart(document.getElementById('chart_div_01'));
		  chart.draw(data, options);
<% 		End If %>
	  }

      function drawChart_02() {
		var ChartX = -1;
		SeriesNames = new Array();
		ChartColors = new Array();

		// Create the data table.
		data = new google.visualization.DataTable();
		data.addColumn('string', 'PhoneLine');

	  <%
		set rs=Server.CreateObject("ADODB.Recordset")
		rs.CursorLocation = 3

		strSQL = "ha_Phone_CallLog_chart_02 ('" & Fltr & "')"
		rs.open strSQL, connCasper10, 3, 3, 4
		
		PrevPhoneLine = ""
		RecCnt = 0
		
		Do until rs.eof
			PhoneLine = rs("PhoneLine")

			If PhoneLine <> PrevPhoneLine Then %>
				ChartX = ChartX + 1;
				data.addRows(1);
				data.setValue(ChartX, 0, '<%=PhoneLine%>');
		<%		PrevPhoneLine = PhoneLine
			End If
		%>
			data.setValue(ChartX, SeriesNbr("<%=rs("FamiliarName")%>", "<%=rs("Color")%>"), <%=rs("CallCount")%>);
		<%	RecCnt = RecCnt + 1
			
			rs.MoveNext
		Loop

		rs.Close
		set rs = nothing
		
		If RecCnt > 0 Then %>
		  // Set chart options
		  var options = {'title':'Total Calls',
						 'isStacked':true,
						 'colors':ChartColors,
						 'hAxis':{showTextEvery:1, slantedText:false, maxAlternation:2, textStyle:{fontSize:13}},
						 'chartArea':{left:50, top:40, width:570, height:400},
						 'backgroundColor':{stroke:"black", strokeWidth:1},
						 'width':870,
						 'height':490
						};

		  // Instantiate and draw our chart, passing in some options.
		  var chart = new google.visualization.ColumnChart(document.getElementById('chart_div_02'));
		  chart.draw(data, options);
<% 		End If %>
	  }

    </script>
<%	End If %>
</head>

<body class="gray_desktop" onkeypress="keytest(event);">
  <table align="center" cellspacing="0" cellpadding="0" class="MainBody">
    <tr>
	  <td valign="top">

		<!--#include virtual="ha_Backoffice/includes/bodyparts/ha_PageHeading.inc"-->
		
	    <form name="frm" action="ha_TelephoneStatistics.asp" method="post">
			<table width="45%" align="center" border="0" cellspacing="0" cellpadding="0">
			  <tr>
				<td />
				<td />
				<td align="left" style="color: black; font-weight: bold;">Start</td>
				<td align="left" style="color: black; font-weight: bold;">End</td>
				<td />
			  </tr>
			  <tr>
				<td width="25%">
					<img src="<%=strPathImages%>blank.gif" width="180" height="1" /
				</td>
				<td width="14%" height="25" align="left" style="color: black; font-weight: bold;" nowrap="nowrap">Date Range:&nbsp;&nbsp;</td>
				<td width="18%" height="25" align="left" nowrap>
					<input name="startDate" id="startDate" type="text" size="9" value="<%=StartDate%>" />
					<a href="javascript:show_calendar('document.frm.startDate', document.frm.startDate.value);">
					<img src="datepick/cal.gif" width="16" height="16" border="0" alt="Calendar" /></a>
				</td>
				<td width="18%" align="left" nowrap>
					<input name="endDate" id="endDate" type="text" size="9" value="<%=EndDate%>" />
					<a href="javascript:show_calendar('document.frm.endDate', document.frm.endDate.value);">
					<img src="datepick/cal.gif" width="16" height="16" border="0" alt="Calendar" /></a>
				</td>
				<td width="25%" align="left" nowrap>
					<a href="javascript:changeDates('<%=DateAdd("d",-1,endDate)%>');" class="greenButton">&lt; Previous</a>
					<a href="javascript:changeDates('<%=date()%>');" class="greenButton">Today</a>
					<a href="javascript:changeDates('<%=DateAdd("d",1,endDate)%>');" class="greenButton">Next &gt;</a>
				</td>
			  </tr>
			  <tr>
				<td />
				<td align="left" height="25" style="color: black; font-weight: bold;">Extension:</td>
				<td colspan="2" align="left">
					<input type="hidden" name="extName" id="extName" />
					<select name="extension" id="extension">
					<option value="-1"></option>
						<%
					set rs=Server.CreateObject("ADODB.Recordset")
					rs.CursorLocation = 3

					strSQL = "ha_Phone_Ext_dropdown"
                    
					rs.open strSQL, connCasper10, 3, 3, 4

					Do until rs.eof
					%>
					<option value="<%=rs("Extension")%>"<%If CStr(rs("Extension"))=CStr(Extension) then%> SELECTED <%End If%>><%=rs("Location")%></option>
					<%	rs.MoveNext
					Loop

					rs.Close
					set rs = nothing
					%>
					</select>
				</td>
				<td />
				<td />
			  </tr>
			  <tr>
				<td colspan="5" height="30" align="center">
					<a href="javascript:clickSubmit();" class="greenButton">&nbsp;Submit&nbsp;</a>
					<input type="hidden" name="submitted" id="submitted" value="Y" />
				</td>
			  </tr>
			</table>
		</form>
		
		<%If Submitted = "Y" Then %>
			<table width="30%" align="center" border="0" cellspacing="0" cellpadding="0">
			<%
				strSQL = "ha_Phone_Statistics_view ('" & Fltr & "')"

				set rs=server.createObject("ADODB.Recordset")
				rs.CursorLocation = 3
                
				rs.open strSQL, connCasper10, 3, 3, 4

				If rs.eof Then %>
				
			  <tr>
				<td height="40" align="center" style="color: black; font-weight: bold;">No records found matching that search criteria.</td>
			  </tr>
			<%
				Else
			%>
			  <tr>
				<td>
				  <table width="100%" align="right" cellspacing="0" cellpadding="4" class="table";>
					<tr bgcolor="#C8E181">
					  <th width="60%">Phone Line</th>
					  <th width="15%"> # Calls</th>
					  <th width="25%">Average<br />Duration</th>
					</tr>
			<%
					Do Until rs.eof
			%>
					<tr>
					  <td align="center" style="border-color: #EEEEEE;">
						<% If rs("Incoming") = "True" Then %>
							<%=rs("PhoneLine") & " -- Incoming calls"%>
						<% Else %>
							<%=rs("PhoneLine") & " -- Outgoing calls"%>
						<% End If %>
					  </td>
					  <td align="right" style="border-color: #EEEEEE;">
						<%=rs("CallCount")%>
					  </td>
					  <td align="center" style="border-color: #EEEEEE;">
						<%=rs("AverageDuration")%>
					  </td>
					</tr>
			<%
						rs.MoveNext
					Loop

					rs.close
					set rs = nothing
			%>
				  </table>
				</td>
			  </tr>
			<%  End If %>
			</table>
		<%End If%>
	  </td>
	</tr>
	<tr>
		<td height="15">&nbsp;</td>
	</tr>
	<tr>
		<td align="center">
		<!--Div that will hold the first chart-->
		<div id="chart_div_01"></div>
		</td>
	</tr>
	<tr>
		<td height="15">&nbsp;</td>
	</tr>
	<tr>
		<td align="center">
		<!--Div that will hold the second chart-->
		<div id="chart_div_02"></div>
		</td>
	</tr>
	<tr>
		<td height="10">&nbsp;</td>
	</tr>
  </table>
  <!--#include virtual="ha_BackOffice/includes/BodyParts/ha_PageTail.inc"-->
</body>
</html>
