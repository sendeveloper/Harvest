<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/Config.asp"-->

<%
	ColorTab = 4
	PageHeading = "Intern Timeline"
%>

<html>
<head>
    <title>Harvest American Backoffice - Intern Timeline</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

    <link href="<%=strPathIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet">
    <link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>menu/dropdowntabfiles/halfmoontabs.css">

    <script type="text/javascript" src="<%=strPathIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>

    <style type="text/css">
      
			span.subHead
				{
				font-weight: 	bold;
				color:			#336600;
				}
			span.superHead
				{
				font-size:		18pt;
				font-weight:	bold;
				
				}
			.titles
				{
				padding-left:3em;
				}
				
				
			span.newStyleHeading
				{
				font-size: 		11pt;
				font-weight: 	bold;
				color:			#336600;
				}
	</style>

    <script language="javascript" type = "text/javascript">
    </script>
</head>

<body class="gray_desktop">
  <table align="center" cellspacing="0" cellpadding="0" class="MainBody">
    <tr>
	  <td>
		
	
	<!--#include virtual="ha_Backoffice/includes/bodyparts/ha_PageHeading.inc"-->

		
		&nbsp&nbsp<form name=frmSemester>
			<select name=mySelect onChange="location.href=frmSemester.mySelect.options[frmSemester.mySelect.selectedIndex].value">
				<option value=""> Select a Semester
				<option value="#summer2012">Summer 2012
				<option value="#fall2012">Fall 2012
				<option value="#spring2013">Spring 2013
				<option value="#summer2013">Summer 2013
				<option value="#spring2014">Spring 2014
				<option value="#summer2014">Summer 2014
				<option value="#fall2014">Fall 2014
		</select>
	</form>
		<br>
		
		<!--
			<b>&nbsp&nbsp<u>Fall 2017</u></b><br>
			<b>&nbsp&nbsp<u>Summer 2017</u></b><br>
			<b>&nbsp&nbsp<u>Spring 2017</u></b><br>
			<b>&nbsp&nbsp<u>Fall 2016</u></b><br>
			<b>&nbsp&nbsp<u>Summer 2016</u></b><br>
			<b>&nbsp&nbsp<u>Spring 2016</u></b><br>
			<b>&nbsp&nbsp<u>Fall 2015</u></b><br>
			<b>&nbsp&nbsp<u>Summer 2015</u></b><br>
			<b>&nbsp&nbsp<u>Spring 2015</u></b><br>
			<b>&nbsp&nbsp<u>Fall 2014</u></b><br> 
		-->
		&nbsp&nbsp<a name="fall2014"><span class="newStyleHeading">Fall 2014</span></a><br>
		&nbsp&nbsp<br>
		&nbsp&nbsp<b>Abhi</b> (co-op) - Coder - SU Remote
		&nbsp&nbsp<ul 
			<ul>
				<li><b>Raffle Ticket Software</b></li>
			</ul>
		<br>&nbsp&nbsp<b>Kanchan</b> (co-op) - Database - SU Remote
		&nbsp&nbsp<ul
		<ul>
		</ul>
		<br>
		&nbsp&nbsp<b>Swati</b>  (co-op) - Coder - SU Remote
		&nbsp&nbsp<ul
			<ul>
			</ul>
		
		<hr width="98%">
		&nbsp&nbsp<a name="summer2014"><span class="newStyleHeading">Summer 2014</span></a>
		<br>&nbsp&nbsp
		<br>&nbsp&nbsp<b>Abhi</b> (co-op) - Coder - Camden Apartment
		&nbsp&nbsp<ul 
			<ul>
				<li><b>Raffle Ticket Software</b></li>
			</ul>
		<br>&nbsp&nbsp<b>Kanchan</b> (co-op) - Database - Camden Apartment
		&nbsp&nbsp<ul
		<ul>
		</ul>
		<br>
		&nbsp&nbsp<b>Swati</b>  (co-op) - Coder - Camden Apartment
		&nbsp&nbsp<ul
			<ul>
				<li>API</li>
				<li>Database Link</li>
				<li>MySQL</li>
			</ul>
		<hr width="98%">
		&nbsp&nbsp<a name="spring2014"><span class="newStyleHeading">Spring 2014</span></a><br>
		
		&nbsp&nbsp<br>
		
		&nbsp&nbsp<b>Pramoth</b> (co-op) - Coder - SU Remote
		&nbsp&nbsp<ul>
				<li>Fixes to Yogesh's PinPoint structure</li>
			</ul>
			<br>
		&nbsp&nbsp<b>Nevidita</b> (co-op) - Server Maintenance - SU Remote
		&nbsp&nbsp<ul>
				<li>Server backups</li>
			</ul>
	<hr width="98%">	
	&nbsp&nbsp<a name="fall2013"><span class="newStyleHeading">Fall 2013</span></a><br><br>
	&nbsp
	<b><i>&nbsp&nbsp No Interns This Semester</i></b>
	<hr width="98%">
	
	&nbsp&nbsp<a name="summer2013"><span class="newStyleHeading">Summer 2013</span></a><br>
		&nbsp<br>
		&nbsp&nbsp<b>Pramoth</b> (co-op) - Coder - Camden Apartment
	&nbsp&nbsp<ul>
				<li>Widget update to use API</li>
				<li>Z2T data origination test</li>
			</ul>
			<br>
		&nbsp&nbsp<b>Sashi</b> (co-op) - Coder - Camden Apartment <i>(quit after first week)</i><br>
	<hr width="98%">  
	&nbsp&nbsp<a name="spring2013"><span class="newStyleHeading">Spring 2013</span></a><br>
	&nbsp<br>		
		&nbsp&nbsp<b>Jo</b> (co-op) - Coder - SU Remote, part-time<br>
		<br>
		&nbsp&nbsp<b>Yogesh</b> (co-op) - Database - SU Remote, part-time<br>
	
	
	<hr width="98%">
	&nbsp&nbsp<a name="fall2012"><span class="newStyleHeading">Fall 2012</span></a>
		<br>&nbsp
		<br>&nbsp&nbsp<b>Deepak</b> (co-op) - Software Designer - SU Remote
	&nbsp&nbsp<ul>
				<li>Raffle Ticket Software Universal Version - <i><b>failed</b></i></li>
			</ul>
		
		<br>&nbsp&nbsp<b>Nivedita</b> (co-op) - Server Maintenance - SU Remote
	&nbsp&nbsp<ul>
				<li>Server backups</li>
			</ul>
	
	
	<hr width="98%"> 
	&nbsp&nbsp<a name="summer2012"><span class="newStyleHeading">Summer 2012</span></a>
	<br>&nbsp
	<br>&nbsp&nbsp<b>Som</b> - Database -<i> Never Started</i><br>
		
		<br>&nbsp&nbsp<b>Jo</b> (co-op) - Coder - Camden Apartment
	&nbsp&nbsp<ul>
				<li>Z2T Widget</li>
				<li>Pinpoint Web Interface</li>
				<li>Cybersource Invoicing</li>
				<li>Ticket Manager framework</li>
			</ul>
				<br>&nbsp&nbsp<b>Yogesh</b> Direct Hire (Recommended by Jo) - Database - Camden Apartment
	&nbsp&nbsp<ul>
				<li>Pinpoint API fixes</li>
			</ul>
<%
	For i = 1 to 25
		Response.Write "<tr><td>&nbsp;</td></tr>"
	Next
%>
	  </td>
	</tr>
  </table>
 </body>
</html>

