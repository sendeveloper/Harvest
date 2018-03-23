<%
  Session("Redirect") = ""
  ColorTab = 1
  PageHeading = "Servers Remote Management"
%>
<!--#include virtual="/includes/connection.asp"--> <!-- ha_backoffico/includes/config will override paths -->
<!--#include virtual="/ha_backoffice/includes/Config.asp"-->
<!DOCTYPE html>
<html>
<head>
  <title>Harvest American Backoffice - <%=PageHeading%></title>
  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"/>

  <link href="<%=strPathBOIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet"/>
  <link rel="stylesheet" type="text/css" href="<%=strPathBOIncludes%>menu/dropdowntabfiles/halfmoontabs.css"/>
  <script type="text/javascript" src="<%=strPathBOIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>
  <script type="text/javascript" src="<%=strPathBOIncludes%>haFunctions.js" language="javascript"></script>
  
  <style type="text/css">

    span.plain
		{
		font-family: courier;
		}
  </style>
  

  <script language="javascript" type = "text/javascript"> 
	  
  </script>
  
</head>

<body class="gray_desktop">
  <div class="MainBody center">
    <!--#include virtual="/ha_Backoffice/includes/bodyparts/ha_PageHeading.inc"-->
    <div style="width: 85%; margin: 20 0 20 100; font-size: 11pt;">

	  The Remote Management Console (RMC) is used to access our servers through a priviledged port (back door).  From here
	  we can check the servers status and perform power recycles (re-boots).<br><br>
	  
	  1.) Using Remote Desktop go into Casper10 (if Casper10 is down use Camden3).<br><br>
	  
	  2.) Open a DOS command window by typing <span class="plain">cmd</span> in the "Search Programs and Files" box at the lower portion of the start menu.<br><br>
	  
	  3.) Move to the location of the ipmish program by entering . . . <br>
	          &emsp;&emsp;&emsp;<span class="plain">cd \dell\bmc</span><br><br>
			  
	  4.) Start ipmish<br>
	          &emsp;&emsp;&emsp;<span class="plain">ipmish -interactive</span><br><br>
			  
	  5.) Depending on the server you wish to enter, select from the following connection lines . . .<br>
	          &emsp;&emsp;&emsp;Casper06:  <span class="plain">connect -ip 66.119.55.119 -k a01378e4aa19418aa59f2e5ca1f3b2a9dc97d8d8 -u root -p get2it</span><br>
	          &emsp;&emsp;&emsp;Casper07:  <span class="plain">connect -ip 66.119.55.120 -k a01378e4aa19418aa59f2e5ca1f3b2a9dc97d8d8 -u root -p get2it</span><br>
	          &emsp;&emsp;&emsp;Casper08:  <span class="plain">connect -ip 66.119.55.121 -k a01378e4aa19418aa59f2e5ca1f3b2a9dc97d8d8 -u root -p get2it</span><br>
	          &emsp;&emsp;&emsp;Casper09:  <span class="plain">connect -ip 66.119.55.122 -k a01378e4aa19418aa59f2e5ca1f3b2a9dc97d8d8 -u root -p get2it</span><br>
	          &emsp;&emsp;&emsp;Casper10:  <span class="plain">connect -ip 66.119.55.123 -k a01378e4aa19418aa59f2e5ca1f3b2a9dc97d8d8 -u root -p get2it</span><br>
	          &emsp;&emsp;&emsp;Philly01:  <span class="plain">connect -ip 208.88.49.23 -u root -p get2it</span><br>
	          &emsp;&emsp;&emsp;Philly02:  <span class="plain">connect -ip 208.88.49.24 -u root -p get2it</span><br>
	          &emsp;&emsp;&emsp;Philly03:  <span class="plain">connect -ip 208.88.49.25 -u root -p get2it</span><br>
	          &emsp;&emsp;&emsp;Philly04:  <span class="plain">connect -ip 208.88.49.26 -u root -p get2it</span><br>
	          &emsp;&emsp;&emsp;Philly05:  <span class="plain">connect -ip 208.88.49.27 -u root -p get2it</span><br><br>
	  
	  6.) Assuming you are now connected, you can enter any IPMISH command<br>
	          &emsp;&emsp;&emsp;<span class="plain">power status</span>	Tells you if the server is on<br>
			  &emsp;&emsp;&emsp;<span class="plain">power cycle</span> Performs a gentle reboot<br><br>
	  
    </div>
  </div>
</body>
</html>

