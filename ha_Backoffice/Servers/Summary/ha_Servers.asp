<!DOCTYPE html>
<html>
<head>

<!--#include virtual="/includes/connection.asp"-->
<!--#include virtual="/ha_backoffice/includes/Config.asp"-->

<%
  Session("Redirect") = ""
  ColorTab = 1
  PageHeading = "Servers Summary"
%>

  <title>Harvest American Backoffice - <%=PageHeading%></title>
  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"/>

  <link href="<%=strPathIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet">  
  <link rel="stylesheet" type="text/css" href="<%=strPathBOIncludes%>menu/dropdowntabfiles/halfmoontabs.css"/>
  
  <script type="text/javascript" src="<%=strPathBOIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>
  <script type="text/javascript" src="<%=strPathBOIncludes%>haFunctions.js" language="javascript"></script>
  <style type="text/css">

    div.columns
    {
      width: 100%;
      display: block;
      margin-bottom: 2em;
    }

    div.column
    {
      display: inline-block;
      vertical-align: top;
      margin-left: auto;
      margin-right: auto;
      width: 30%;
    }

    div
    {
      font-family: Arial, Helvetica, sans-serif;
      font-size: 12px;
      font-style: normal;
      line-height: normal;
      font-weight: normal;
      font-variant: normal;
      text-transform: none;
      color: #000000;
      text-decoration: none;}

    div.block-div
    {
      display: block;
    }

    div.inline-block-div
    {
      display: inline-block;
      font-family: Arial, Helvetica, sans-serif;
        font-size: 12px;
        font-style: normal;
        line-height: normal;
        font-weight: normal;
        font-variant: normal;
        text-transform: none;
        color: #000000;
        text-decoration: none;}

    .server-farm-title
		{
        font-size: 		12pt;
		font-weight:	bold;
        text-align: 	left;
		width: 			95%;
		margin-top:		20px;
		}
		
    .server-farm-title a
		{
		color:			black;
		}
		
    div.server-farm
		{
        vertical-align: top;
        border: 2px solid black;
        background-color: #FFFFCC;
        width: 100%;
        height: auto;
        display: inline-block;
		}

    div.server-farm a
		{
		color:			blue;
		}
		
    div.server
    {
        border: 5px inset black;
        background-color: #DDDDDD;
        width: 90%;
        display: block;
        margin-left: auto;
        margin-right: auto;
        margin-top: 1em;
        margin-bottom: 1em;}

        div.server *
        {
            line-height: 200%;}

    div.server-status
    {
      margin-left: 0.5em;
      margin-right:0.5em;
      width: 95%;
    }

    div.server-name
    {
        font-size: 14px;
        font-weight: bold;
        width: 50%;
        text-align: left;
        /*margin-left: .5em;*/
        margin-left: 0em;
        margin-right: auto;
        display: inline-block;
    }

    div.test
    {
      text-align: right;
      display: inline-block;
      margin-left: auto;
      margin-right: 0em;
      width: 49%;
      color: black;

    }

    div.test > div
    {
      background: lightgray;
      display: inline-block;
      margin-right: .5em;
      cursor: pointer;
    }

    div.test.success
    {
      background-color: green;
      color: white;
    }

    div.test.failed
    {

    }

    div.server-type
    {
        font-size: 12px;
        text-align: center;}

    div.server-ip
    {
        font-size: 12px;
        text-align: center;}

        div.center
        {
            text-align: center;}

  </style>
  <script language="javascript" type = "text/javascript">
    function clickProperties(ServerID)
    {
      var URL = '<%=strPathServers%>Summary/ha_ServerProperties.asp' +
                '?ServerID=' + ServerID;
      openPopUp(URL,'90', '600');
    }

    function clickPing(ServerID)
    {
      var URL = '<%=strPathServers%>Summary/ha_ServerPing.asp' +
                '?ServerID=' + ServerID;
      openPopUp(URL);
    }

    function clickPopup(n)
    {
      var URL = '<%=strPathServers%>Summary/' + n + '.asp'
      openPopUp(URL);
    }
    </script>
</head>

<body class="gray_desktop">
  <div class="MainBody center">
    <!--#include virtual="/ha_Backoffice/includes/bodyparts/ha_PageHeading.inc"-->

    <div class="columns">
	  <span style="text-align: right; display: block;">
        <a style="display: inline-block; padding-right: 2em;" href="/ha_BackOffice/Servers/Summary/ha_Servers_listview.asp">List View</a>
	  </span>
      <div class="column column-1">

        <div class="server-farm-title">
          <a href="http://www.mwtn.net/datacenter/offerings/overview" target = "_new">
		    Mountain West - Casper, WY</a>
          <div class="server-farm">
			<%=DisplayServer("Casper06", "Zip2Tax (West)")%>
			<%=DisplayServer("Casper07", "Web Server/SQL Server")%>
			<%=DisplayServer("Casper08", "Web Server/SQL Server")%>
			<%=DisplayServer("Casper09", "E-mail - Zip2Tax Development")%>
			<%=DisplayServer("Casper10", "Backoffice Operations")%>	
			
            <span style="width: 100%; text-align: right; display: block;">
              <a href="javascript:clickPopup('ha_Servers_Network_Casper')" style="font-size: 10pt; padding-right: 15px;">Network Diagram</a>
              <br>&nbsp;
            </span>
			
          </div><!-- serverfarm -->
        </div><!-- server-farm-title -->
		
      </div><!-- column-1 -->
	  

      <div class="column column-2">
	  
        <div class="server-farm-title">
          <a href="http://www.HarvestAmerican.com" target = "_new">
		    Harvest American, Camden, NY</a>
		  &nbsp;&nbsp;
          <a href="javascript:clickPopup('ha_Servers_Image_Camden')" style="font-size: 10pt;">Image</a>
          <div class="server-farm">
			<%=DisplayServer("Camden01", "SQL Database")%>
			<%=DisplayServer("Camden02", "FTP Server")%>
			<%=DisplayServer("Camden03", "Phone Log - Backup Server - File Server")%>
			<%=DisplayServer("Camden04", "Domain Controller - Active Directory")%>
			<%=DisplayServer("Camden05", "Telephone Interface")%>
            <span style="width: 100%; text-align: right; display: block;">
			  <span style="font-size: 10pt; font-weight: bold;">Networks:&nbsp;&nbsp;</span>
              <a href="javascript:clickPopup('ha_Servers_Network_Camden_Cable')" style="font-size: 10pt; padding-right: 15px;">Cable</a>
              <a href="javascript:clickPopup('ha_Servers_Network_Camden_DSL')" style="font-size: 10pt; padding-right: 15px;">DSL</a>
              <a href="javascript:clickPopup('ha_Servers_Network_Camden_T1')" style="font-size: 10pt; padding-right: 15px;">LAN</a>
              <br>&nbsp;
            </span>
          </div><!-- server-farm -->
        </div><!-- server-farm-title -->

        <div class="server-farm-title">
          <a href="http://www.dfarns.com" target = "_new">
		    Dave's House, McConnellsville, NY</a>
          <div class="server-farm margin: .5em; padding: .5em;">
			<%=DisplayServer("McVille02", "File Server - Backup Server")%>
            <span style="width: 100%; text-align: right; display: block;">
              <a href="javascript:clickPopup('ha_Servers_Network_McVille')" style="font-size: 10pt; padding-right: 15px;">Network Diagram</a>
              <br>&nbsp;
            </span>
          </div><!-- server-farm -->
        </div><!-- server-farm-title -->
		
        <div class="server-farm-title">
          <a href="http://www.GoDaddy.com" target = "_new">
			GoDaddy - Scottsdale, AZ</a>
          <div class="server-farm">
			<%=DisplayServer("Barley", "Web Server - SQL Database")%>
			<%=DisplayServer("Barley2", "SQL Database - MySQL Database")%>
          </div><!-- serverfarm -->
        </div><!-- server-farm-title -->
		

        <div class="server-farm-title">
          <a href="http://www.ActiveServers.com" target = "_new">
			Active Servers - Spokane, WA</a>
          <div class="server-farm">
			<%=DisplayServer("Viking", "Web Server")%>
          </div><!-- server-farm -->
        </div><!-- server-farm-title -->
      </div><!-- column-2 -->

      <div class="column column-3">
        <div class="server-farm-title">
          <a href="javascript:alert('Quonix Networks\nAttn: John Von Essen\n2401 Locust St\nPhila, PA 19103\njohn@quonix.net\n215-563-1831')" target = "_new">
		    Quonix Networks - Philadelphia, PA</a>
          <div class="server-farm">

			<%=DisplayServer("Philly01", "Gatekeeper - Downloader")%>
			<%=DisplayServer("Philly02", "Zip2Tax SQL Server #1")%>
			<%=DisplayServer("Philly03", "Z2T Web Server - Website - Badges - table downloads")%>
			<%=DisplayServer("Philly04", "Zip2Tax SQL Server #2")%>
			<%=DisplayServer("Philly05", "Export - Research - PinPoint Dev - z2t Backoffice")%>		

            <span style="width: 100%; text-align: right; display: block;">
              <a href="javascript:clickPopup('ha_Servers_Network_Philly')" style="font-size: 10pt; padding-right: 15px;">Network Diagram</a>
              <br>&nbsp;
            </span>
          </div><!-- server-farm -->
        </div><!-- server-farm-title -->
		
        <div class="server-farm-title">
          <a href="http://www.first-colo.net" target = "_new">
		    First-Colo  - Frankfurt, Germany</a>
          <div class="server-farm">

			<%=DisplayServer("Frank01", "Domain Controller - Active Directory")%>
			<%=DisplayServer("Frank02", "Zip2Tax Production")%>
			<%=DisplayServer("Frank03", "Zip2Tax Production")%>

            <span style="width: 100%; text-align: right; display: block;">
              <a href="javascript:clickPopup('ha_Servers_Network_Frank')" style="font-size: 10pt; padding-right: 15px;">Network Diagram</a>
              <br>&nbsp;
            </span>
          </div><!-- server-farm -->
        </div><!-- server-farm-title -->
      </div><!-- column-3 -->
	  
    </div><!-- columns -->
  </div><!-- MainBody -->
  <!--#include virtual="ha_BackOffice/includes/BodyParts/ha_PageTail.inc"-->
  
  <script src="/ha_Backoffice/includes/lib.js"></script>
  <script>
    //get(".server-name").forEach(function(e){e.forEach(function(e){alert(e.length)})});
    get(".server-name").forEach(function(e){e.forEach(function(e){
      return e.parentNode.insertBefore(node("div", ["class","test"], node("div", [], text("ping")), node("div", [], text("iis")), node("div", [], text("sql"))), e.nextSibling);});});
    //get(".test div").forEach(function(e){e.forEach(function(e){listen(e, "click", function(e){ajax(protocol + "://" + server + "/includes/wsh.asp?command=" + command,
    get(".test div").forEach(function(e){e.forEach(function(e){listen(e, "click", function(e){
      e.target.style.backgroundColor = "gray";
      //alert("ping -n 1 " + e.target.parentNode.previousSibling.innerHTML.trim());
      var ping = ajax("/ha_BackOffice/includes/wsh.asp?command=ping -n 1 " + e.target.parentNode.previousSibling.innerHTML.trim() + ".harvestamerican.net",
      //var ping = ajax("/ha_BackOffice/includes/wsh.asp?command=ping -n 1 localhost",
        function(e){
          var ms = extractPingAverage(e);
          this.node.title = ms;
          //alert(this.node.title);
          this.node.style.backgroundColor = (ms.length <= 25) ? "green" : "red";},
        function(e){
          this.node.title = e;
          this.node.style.backgroundColor = "red";});
      ping.node = e.target;
      ping.get();})})});

  function extractPingAverage(ping) {
    return ensurePattern(ping, /Average = ([0-9]+)ms/, 
      function(e) {return "ping-average: " + e[1] + "ms";},
      function(e) {return "ping-failed: " + ping;});}
  
  function ensurePattern(text, pattern, fn, err) {
    matches = text.match(pattern);
    return matches ?
      fn(matches) :
      err ? err(text, pattern) : "Pattern \"" + pattern.toString() + "\" was not found in \"" + text + "\".";}

  </script>
</body>
</html>

<%

Function DisplayServer(servername, description)

	If Len(description) > 35 Then
		description = "<span style='font-size: 8pt;'>" & description & "</span>"
	End If

	DisplayServer = vbCrLf & _
		"            <div class='server'>" & vbCrLf & _
		"              <div class='server-status'>" & vbCrLf & _
		"                <div class='server-name'>" & vbCrLf & _
		"                  " & servername & vbCrLf & _
		"                </div><!-- server-name -->" & vbCrLf & _
		"              </div><!-- server-status -->" & vbCrLf & _
		"              <div class='server-type'>" & vbCrLf & _
		"                " & description & vbCrLf & _
		"              </div><!-- server-type -->" & vbCrLf & _
		"              <div class='server-ip'>" & vbCrLf & _
		"                " & LinkLine(servername) & vbCrLf & _
		"              </div><!-- server-ip -->" & vbCrLf & _
		"            </div><!-- server -->" & vbCrLf
	
End Function

Function LinkLine(servername)

    LinkLine = ""

    set rs=Server.CreateObject("ADODB.Recordset")

    'Read Properties
    sql = "SELECT Count(*) c from ha_Types_repl WHERE Class = '" & servername & " - Properties'"
    rs.open sql, connCasper10, 1, 2
    if not rs.eof then
        if cInt(rs("c") > 0) then
            LinkLine = LinkLine & "<a href=""javascript:clickProperties('" & servername & " - Properties')"" class='Button50'>Properties</a>&nbsp;"
		Else
			LinkLine = "&nbsp;"		
        end if
    end if
    rs.close

    'Read IP
    sql = "SELECT Value from ha_Types_repl WHERE Class = '" & servername & " - Properties' AND Description = 'IP'"
    rs.open sql, connCasper10, 1, 2
    if not rs.eof then
        LinkLine = LinkLine & "<a href=""javascript:clickPing('" & rs("Value") & "')"" class='Button50'>Ping IP</a>&nbsp;"
    end if
    rs.close

    'Read Domain Name
    sql = "SELECT Value from ha_Types_repl WHERE Class = '" & servername & " - Properties' AND Description = 'Domain Name'"
    rs.open sql, connCasper10, 1, 2
    if not rs.eof then
        LinkLine = LinkLine & "<a href=""javascript:clickPing('" & rs("Value") & "')"" class='Button50'>Ping Domain</a>&nbsp;"
    end if
    rs.close

    'Additional Services
    Select Case servername
      Case "Camden1"
        LinkLine = LinkLine & "<a href=""" & strPathServers & "ha_Services_last.asp"" class='Button50'>Services</a>&nbsp;"
      Case "Camden3"
        LinkLine = LinkLine & "<a href=""" & strPathTelephones & "ha_TelephoneCallLog.asp"" class='Button50'>Phone Log</a>&nbsp;"
    End Select

   If instr(servername, "Philly") Then
     LinkLine = LinkLine & "<a href=""javascript:alert('Server Logs are under construction')"">Logs</a>"
   End If

End Function

%>

