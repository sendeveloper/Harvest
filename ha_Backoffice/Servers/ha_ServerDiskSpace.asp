<!DOCTYPE html>
<!--#include virtual="includes/connection.asp"-->
<!--#include virtual="ha_backoffice/includes/Config.asp"-->
<!--#include virtual="ha_backoffice/includes/FunctionsTimeDifference.inc"-->

<%
	ColorTab = 1
	PageHeading = "Server Disk Space"

    Function GBConv(MB)
      GBConv = MB / 1024
    End Function
	
	ToolTip1 = "<span class='ToolTipBox'>Click box anywhere to close.</span>"
	ToolTip2 = "<span class='ToolTipBox'>Click to view larger image.</span>"

%>

<html>
<head>
    <title>Harvest American Backoffice - <%=PageHeading%></title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />

    <link href="<%=strPathIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet" />
    <link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>menu/dropdowntabfiles/halfmoontabs.css" />

    <script type="text/javascript" src="<%=strPathIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>

    <style type="text/css">
		th
			{
			border-bottom:	1px solid black;
			}

		th.Heading1
			{
			font-size:		14px;
			text-align:		center;
			border-bottom:	1px solid black;
			}
	  
      .num
      {
       text-align: right;   
      }
      .head
      {
          color: White;
      }
    </style>
    <script type="text/javascript" src="https://www.google.com/jsapi"></script> 

    <script type="text/javascript">
        // Load the Visualization API and the column chart package.
        google.load('visualization', '1.0', { packages: ['corechart'] });

        google.setOnLoadCallBack(drawChart);

        function drawChart(srvname, divid, drvltr, free, total, w, h) {
            var used = total - free;
            var data = google.visualization.arrayToDataTable([
            
            ['Type', 'Percent'],
            
            ['Space Free', free],
            
            ['Space Used', used]

            ]);
            var options = { title: drvltr + ' Disk Space in GB for Server ' + srvname,
                            width: w,
                            height: h };
            
            // Create and draw the visualization.
            var chart = new google.visualization.PieChart(document.getElementById(divid));
            chart.draw(data, options);

            //google.visualization.events.addListener(chart, 'select', selectHandler);

            //function selectHandler(e) {
            //    chart.clearChart(); 
            //}
            document.getElementById('diskspacechart').style.display = 'block';
        }
    </script>
</head>

<body class="gray_desktop">
  <table align="center" cellspacing="0" cellpadding="0" class="MainBody">
    <tr>
      <td>
		<!--#include virtual="ha_Backoffice/includes/bodyparts/ha_PageHeading.inc"-->
	
		<table width="90%" border="0" cellspacing="4" cellpadding="4" align="center">
		
		  <tr><td>&nbsp;</td></tr>
		
			  <tr>
				<th width="20%" align="left">Server Name</th>
				<th width="40%" align="center">C: Drive</th>
				<th width="40%" align="center">D: Drive</th>
			  </tr>
			  
<%
		recordCount = 0

		set rs=server.createObject("ADODB.Recordset")
		SQL = "ha_Servers_DiskSpace_list"
		rs.Open SQL, connCasper10, 3, 3, 4

		if not RS.EOF then
			do while not rs.eof
				recordCount = recordCount + 1
				
				'Dividing Separator
				If SeparatorRows = "" Then
					'Defaults
					SeparatorRows = 3
					SeparatorWidth = "3px"
					SeparatorColor = "#D8F0E3"
				End If

				If recordCount mod SeparatorRows = 0 Then
					detailStyle = "border-bottom: " + SeparatorWidth + " solid " +SeparatorColor + ";"
				Else
					detailStyle = "border: 0;"
				End If
%>

		  <tr>
			<td style="<%=detailStyle%>">
			  <%=rs("ServerName")%>
			</td>
			
			<td style="<%=detailStyle%>">
			  <table width="100%" border="0" cellspacing="1" cellpadding="1">
			    <tr valign="top">
				  <td width="35%" align="center">
					<%=rs("DisplayC")%>
				  </td>
				  <td width="35%" align="center">
<%
              If rs("MBTotalC") > 0 Then
                response.write " " & FormatPercent(rs("PercentFreeC"),1)  & " Free"
			  Else
			    response.write "&nbsp;"
			  End If
%>
				  </td>
				  <td width="30%" align="center">
<%
              If rs("MBTotalC") > 0 Then
%>
			    <a href="javascript:void(0)" onclick="drawChart('<%=rs("ServerName")%>','container','C:',<%=FormatNumber(GBConv(rs("MBFreeC")),2)%>, <%=FormatNumber(GBConv(rs("MBTotalC")),2)%>, 500, 300)">
				  <div id="smallpie<%=recordCount%>a">

					<script type="text/javascript">
                      document.write(drawChart('<%=rs("ServerName")%>','smallpie<%=recordCount%>a','C:', <% =FormatNumber(GBConv(rs("MBFreeC")),2) %>, <% =FormatNumber(GBConv(rs("MBTotalC")),2) %>,45,24));
					</script>

                  </div></a>
<%
			  Else
			    response.write "&nbsp;"
			  End If
%>
				  </td>
				</tr>
			  </table>
			</td>
			
			<td style="<%=detailStyle%>">
			  <table width="100%" border="0" cellspacing="2" cellpadding="2">
			    <tr>
				  <td width="35%" align="center">
					<%=rs("DisplayD")%>
				  </td>
				  <td width="35%" align="center">
<%
              If rs("MBTotalD") > 0 Then
                response.write " " & FormatPercent(rs("PercentFreeD"),1)  & " Free"
			  Else
			    response.write "&nbsp;"
			  End If
%>
				  </td>
				  <td width="30%" align="center">
<%
              If rs("MBTotalD") > 0 Then
%>
			    <a href="javascript:void(0)" onclick="drawChart('<%=rs("ServerName")%>','container','D:',<%=FormatNumber(GBConv(rs("MBFreeD")),2)%>, <%=FormatNumber(GBConv(rs("MBTotalD")),2)%>, 500, 300)">
				  <div id="smallpie<%=recordCount%>b">

					<script type="text/javascript">
                      document.write(drawChart('<%=rs("ServerName")%>','smallpie<%=recordCount%>b','D:', <% =FormatNumber(GBConv(rs("MBFreeD")),2) %>, <% =FormatNumber(GBConv(rs("MBTotalD")),2) %>,45,24));
					</script>

                  </div></a>
<%
			  Else
			    response.write "&nbsp;"
			  End If
%>
				  </td>
				</tr>
			  </table>
			</td>
		  </tr>

<%
				AsOf = rs("AsOf")
				rs.MoveNext
			loop
		end if
		
		rs.close
%>

		  <tr>
			<th class="Heading1" colspan="3">&nbsp;</th>
		  </tr>
		  			  
		  <tr>
		    <td colspan="3">
			  Last Updated: <%=AsOf%> -- <%=TimeDifference(AsOf, Now)%>
			</td>
		  </tr>

		  <tr><td>&nbsp;</td></tr>
		  
		</table>

      </td>
    </tr>
		

	<!--Popup Box-->
    <div id="diskspacechart"
		style="position: absolute; top: 30%; left: 10%; border: 2px solid black; z-index: 5;" 
		onclick="this.style.display = 'none'">
      <a href="javascript:void(0)" onclick="document.getElementById('container').innerHTML = '';this.style.display = 'none';" 
		style="display: none; position: absolute; margin: auto;">X</a>
      <div id="container" style="width: 95%"> </div>
	</div>
		
  </table>
  
  <!--#include virtual="ha_BackOffice/includes/BodyParts/ha_PageTail.inc"-->
  
</body>
</html>

