<!DOCTYPE html>

<!--#include virtual="includes/connection.asp"-->

<html>
<head>
    <title>Dave's Menu</title>
	<meta charset="UTF-8">
	
    <script type="text/javascript" src="/includes/haFunctions.js"></script>

    <style type="text/css">

.inner-table{

	border:0px;
		border-right:ridge 1px #999	;
		border-radius:0px;
	font-family:Verdana, Geneva, sans-serif;
	}
	table{
font-family:Verdana, Geneva, sans-serif;
		border-radius: 38px;
		border: #616563 1px solid;
		background-color: #f5f5e7;


		}
	td
		{
		font-family:Verdana, Geneva, sans-serif;
		font-size: 12px; 
		font-style: normal; 
		line-height: normal; 
		font-weight: normal; 
		font-variant: normal; 
		text-transform: none;
		text-decoration: none;
		color: #336600;

		}

	th
		{
		font-size: 14px;
		}

	h1, h2, h3, h4, h5, h6 
		{
		color: #336600;
		}

	hr
		{
		color: #BED985;
		}

	a
		{
		font-size: 14px;
		color: #336600;
		text-decoration: none;
		font-family:Verdana, Geneva, sans-serif;
		font-weight:bold;
		}

	a:hover 
		{
		color: #000000;
		text-decoration: underline;
		}

	a.small
		{
		font-size: 14px;
		padding-left: 5px;
		font-weight:400
		}
.no-style{
border:0px;	
font-family:Verdana, Geneva, sans-serif;
	}
body{font-family:Verdana, Geneva, sans-serif;}
    </style>

<script type="text/javascript">

    function clickStatusChange()
        {
        var URL = '<%=strPathBase%>/menu/StaffStatusEdit.asp' +
            '?EmployeeID=<%=Session("ha_EmployeeID")%>';
        openPopUp(URL);
        }
</script>

</head>

<body>

<table width="100%"  cellspacing="5" cellpadding="5">
  <tr width="100%">
    <td colspan="5" align="center" style="font-size:16px;font-weight:bold; ">  
      Dave's Menu
    </td>
  </tr>

  <tr valign="top">

    <td width="20%" align="center">
      <table width="100%" border="0" cellspacing="5" cellpadding="5" class="inner-table">
        <tr>
          <th>
            Orders
          </th>
        </tr>
        <tr>
          <td>
            <a href="https://ebc.cybersource.com/ebc/login/Login.do">CyberSource</a>
          </td>
        </tr>
        <tr>
          <td>
            <a href="http://www.snscart.com/login.htm">Shopping Cart</a>
          </td>
        </tr>
        <tr>
          <td>
            <a href="http://www.number-it.com">Number-it</a>
            <a href="http://www.number-it.com/Home/BackOffice/boLogin.asp?lname=<%=Session("ha_login")%>&pass=<%=Session("ha_password")%>"
              class="small">Backoffice</a>
          </td>
        </tr>
        <tr>
          <td>
            <a href="http://www.amazon.com/gp/seller-account/management/your-account.html/002-3179982-7492809">
              Amazon Seller Account</a>
          </td>
        </tr>


        <tr>
          <th>
            <br>Projects
          </th>
        </tr>
        <tr>
          <td>
            <a href="https://teambox.com/login">TeamBox</a>
          </td>
        </tr>
        <tr>
          <td>
            <a href="https://www.guru.com/login.aspx">Guru.com</a>
          </td>
        </tr>

        <tr>
          <th>
            <br>Your Web Vendors
          </th>
        </tr>
        <tr>
          <td>
            <a href="https://adwords.google.com/select/Login?sourceid=awo&subid=ww-en-et-ads1_e&hl=en_us">
              Google AdWords</a>
            <a href="https://www.google.com/webmasters/tools/siteoverview?hl=en&msg=2&pli=1" class="small">
              Sitemaps</a>
            <a href="https://www.google.com/webmasters/tools/dashboard?hl=en&siteUrl=http://www.zip2tax.com/" class="small">
              Webmaster Tools</a>
          </td>
        </tr>
        <tr>
          <td>
            <a href="https://adcenter.microsoft.com">
              MSN</a>
          </td>
        </tr>
        <tr>
          <td>
            <a href="https://login22.marketingsolutions.yahoo.com">
              Yahoo</a>
          </td>
        </tr>
        <tr>
          <td>
            <a href="http://ConstantContact.com">Constant Contact</a>
            <a href="Content/Constant Contact Email Procedure.rtf" class="small">How</a>
          </td>
        </tr>

      </table>
    </td>

    <td width="20%" align="center">
      <table width="100%" border="0" cellspacing="5" cellpadding="5" class="inner-table">

        <tr>
          <th>
            Banking
          </th>
        </tr>
        <tr>
          <td>
            <a href="https://www.accessfcu.org/onlineserv/HB/Signon.cgi">Access Federal CU</a>
          </td>
        </tr>
        <tr>
          <td>
            <a href="https://sitekey.bankofamerica.com/sas/signonSetup.do">Bank of America</a>
          </td>
        </tr>
        <tr>
          <td>
            <a href="https://mtb.com">M&T Bank</a>
              <a href="https://www.mtb.com/personal/Pages/Index.aspx" class="small">Personal</a>
              <a href="https://www.mtb.com/business/Pages/BusinessHome.aspx" class="small">Business</a>
          </td>
        </tr>
        <tr>
          <td>
            <a href="http://www.USBank.com/index.html">US Bank</a>
              <a href="http://www.USBank.com/index.html" class="small">Personal</a>
              <a href="http://www.usbank.com/en/SmallBusHome.cfm" class="small">Business</a>
          </td>
        </tr>
        <tr>
          <td>
            <a href="http://www.nbtbank.com">NBT</a>
          </td>
        </tr>
        <tr>
          <td>
            <a href="https://CoinBase.com">CoinBase</a>
          </td>
        </tr>

        <tr>
          <th>
            <br>Bookkeeping
          </th>
        </tr>
        <tr>
          <td>
            <a href="<%=strPathBase%>/ha_BackOffice/Company/ha_timeclockmgmt.asp">
              Employee Timeclock Management</a>
          </td>
        </tr>
        <tr>
          <td>
            <a href="<%=strPathBase%>/projects/ha_projectsandprocedures.asp">
              Project &amp; Procedures List</a>
          </td>
        </tr>

        <tr>
          <th>
            <br>Investments
          </th>
        </tr>
        <tr>
          <td>
            <a href="https://wwws.ameritrade.com/apps/LogIn">
              TD Ameritrade</a>
          </td>
        </tr>
        <tr>
          <td>
            <a href="https://www.fidelity.com/">
              Fidelity</a>
          </td>
        </tr>
        <tr>
          <td>
            <a href="http://finance.yahoo.com/q/cp?s=%5EDJI">
              Dow Components</a>
          </td>
        </tr>
        <tr>
          <td>
            <a href="http://myservices.equifax.com/ESA_dashboard">
              EquiFax</a>
          </td>
        </tr>

      </table>
    </td>



    <td width="20%" align="center">
      <table width="100%" border="0" cellspacing="5" cellpadding="5" class="inner-table">
        <tr>
          <th>
            Your Websites
          </th>
        </tr>

        <tr>
          <td>
            <a href="http://www.CamdenBudgetStorage.com/home/login/logpost.asp?lname=dave&pass=go4it">
              Camden Budget Storage</a>
          </td>
        </tr>

        <tr>
          <td>
            <a href="http://www.dfarns.com">dfarns.com</a>
          </td>
        </tr>

        <tr>
          <td>
            <a href="http://www.EyeOnCamden.com">Eye On Camden</a>
          </td>
        </tr>

        <tr>
          <td>
            <a href="http://www.HarvestAmerican.com">Harvest American</a>
              <a href="https://accounting.HarvestAmerican.info/index.html" class="small">Accounting</a>
              <a href="/ha_backoffice" class="small">Backoffice</a>
              <a href="http://crm.HarvestAmerican.info/index.asp" class="small">CRM</a>
              <a href="http://www.HarvestAmerican.info/ControlPanel/stats/ha_Stats.asp" class="small">Stats</a>
              <a href="http://www.harvestamerican.com/wholesale/tickets/"class="small">Wholesale</a>
              <a href="http://helpdesk.harvestamerican.net" class="small">Help Desk</a>
         </td>
        </tr>

        <tr>
          <td>
            <a href="http://www.number-it.com">Number-it</a>
            <a href="http://www.number-it.com/Home/BackOffice/boLogin.asp?lname=<%=Session("ha_login")%>&pass=<%=Session("ha_password")%>"
              class="small">Backoffice</a>
          </td>
        </tr>

        <tr>
          <td>
            <a href="http://www.NYSThruway.com">NYSThruway.com</a>
          </td>
        </tr>

        <tr>
          <td>
            <a href="http://www.WJ2O.com">WJ2O.com</a>
              <a href="Content/Uploading Logs To WJ2O.rtf" class="small">Log Upload</a>
          </td>
        </tr>

        <tr>
          <td>
            <a href="http://www.ShipHeroes.com">ShipHeroes</a>
              <a href="https://www.shipheroes.com/BackOffice/login/sh_login.asp" class="small">
                Backoffice</a>
          </td>
        </tr>

        <tr>
          <td>
            <a href="http://www.zip2tax.com">Zip2Tax</a>
              <a href="http://info.zip2tax.com/z2t_Backoffice/home/z2t_home.asp" class="small">
                New Backoffice</a>
              <a href="http://legacy.zip2tax.com/backoffice/z2t_Login.asp?lname=davewj2o&pass=get2it" class="small">
                Backoffice</a>
              <a href="http://www.harvestamerican.us/common/reference/z2t_dotnums.txt" class="small">
                State Numbers</a>
              <a href="Content/z2t Monthly Updates Procedure.rtf" class="small">
                Update Process</a>
              <a href="Content/z2t Month End Procedures.rtf" class="small">
                Month End Procedures</a>
          </td>
        </tr>



        <tr>
          <th>
            <br>Servers
          </th>
        </tr>

        <tr>
          <td>
            <a href="/ha_BackOffice/Servers/ha_Servers.asp">Harvest Control Panel</a>
              <a href="<%=strPathBase%>/ControlPanel/cp_Servers.asp?lname=nopassword" class="small">No Password</a>
          </td>
        </tr>
        <tr>
          <td>
            <a href="https://68.178.202.54:8443/login.php3">Plesk on Harvest</a>
          </td>
        </tr>
        <tr>
          <td>
            <a href="http://www.godaddy.com">GoDaddy</a>
              <a class="small" href="https://mya.godaddy.com/?ci=60016">My Account</a>
          </td>
        </tr>
        <tr>
          <td>
            <a href="https://cp.dnsmadeeasy.com/">DNS Made Easy</a>
          </td>
        </tr>
        <tr>
          <td>
            <a href="https://www12.qth.com:2083/frontend/x3/index.html">QTH</a>
              <a class="small" href="https://www12.qth.com:2083/frontend/x3/index.html">Control Panel</a>
              <a class="small" href="https://www12.qth.com:2096/webmaillogout.cgi">Webmail</a>
          </td>
        </tr>
        <tr>
          <td>
            <a href="https://hsp.viux.com/cp/index.cgi/top/zone,home/poa_ok,ok/">Viux</a>
              <a class="small" href="https://hsp.viux.com/cp/index.cgi/top/zone,home/poa_ok,ok/">Control Panel</a>
          </td>
        </tr>
        <tr>
          <td>
            <a href="http://www.geopeeker.com">GeoPeeker</a>
          </td>
        </tr>


      </table>
    </td>

    <td width="20%" align="center">
      <table width="100%" border="0" cellspacing="5" cellpadding="5" class="inner-table">

        <tr>
          <th>
            Clients
          </th>
        </tr>

        <tr>
          <td>
            <a href="http://www.tlls.org">Leukemia</a>
              <a href="Contacts_leuk.asp" class="small">Contacts</a>
              <a href="https://webmail.lls.org" class="small">Mail</a>
              <a href="http://SupportCenter.lls.org" class="small">SupportCenter</a>
              <a href="https://vpn.lls.org/webvpn.html" class="small">VPN</a>
              <a href="https://secure.lls.org/Citrix/MetaFrame/default/default.aspx" class="small">Citrix</a>
              <a href="http://jira.tlls.net/" class="small">JIRA</a>
          </td>
        </tr>

        <tr>
          <td>
            <a href="http://www.gotomypc.com">ccfa Production Server</a>
              <a href="Content/How To Login To ccfa Prod Server.txt" class="small">How to log on to CCFA server</a><a href="Contacts_ccfa.asp" class="small">Contacts</a>
          </td>
        </tr>

        <tr>
          <td>
            <a href="https://www.filecollegeinfo.com">cts Website</a>
              <a href="http://test.filecollegeinfo.com" class="small">Test</a>
              <a href="Contacts_cts.asp" class="small">Contacts</a>
          </td>
        </tr>

        <tr>
          <td>
            <a href="http://africa.harvestamerican.net/index.asp">Mkombozi</a>
          </td>
        </tr>

        <tr>
          <th>
            <br>Reference
          </th>
        </tr>
        <tr>
          <td>
            <a href="Content/TelephoneWiringDiagrams.xls">Telephone Wiring Diagrams</a>
          </td>
        </tr>
        <tr>
          <td>
            <a href="http://www.htmlhelp.com/reference/css/properties.html">WDG CSS Properties</a>
          </td>
        </tr>

        <tr>
          <th>
            <br>Phones
          </th>
        </tr>
        <tr>
          <td>
            <a href="/menu/phones.asp">Call Log</a>
          </td>
	</tr>
	<tr>
	  <td>
	    <table style="width: 100%;" class="no-style">
	      <tr>
		<td style="font-size: 8px;"><a href="https://davewj2o:get2it@24.103.166.106">Dave's</a></td>
		<td style="font-size: 8px;"><a href="https://72.43.236.2/cgi-bin/userLogin.cgi?login=true&portalname=CommonPortal&password_expired=0&auth_key=1964300002&md5_old_pass=&username=davewj2o&password=bfe97b8c6382d6fef20ef3805230ca03">Camden</a></td>
		<td style="font-size: 8px;"><a href="https://lrowlands:get2it@lrowlands.dyndns.info">Holland Patent</a></td>
		<td style="font-size: 8px;"><a href="https://jlowe:get2it@jlowe.dyndns.info">North Carolina</a></td>
	      </tr>
	    </table>
	  </td>
        </tr>

        <tr>
          <th>
            <br>Comm100
          </th>
        </tr>
        <tr>
	  <td>
	    <table style="width: 100%;" class="no-style">
	      <tr>
		<td style="font-size: 8px;"><a href="http://hosted.comm100.com/Tickets/ticket.aspx?siteId=49685">Tickets</a></td>
		<td style="font-size: 8px;"><a href="http://hosted.comm100.com/LiveChat/VisitorMonitor.aspx?siteId=49685">Chat</a></td>
		<td style="font-size: 8px;"><a href="http://hosted.comm100.com/KnowledgeBase/Category.aspx?Id=191&siteId=49685">Knowledge Base</a><br></td>
	      </tr>
	    </table>
	  </td>
        </tr>

      </table>
    </td>

	<td width="20%" align="center">
      <table width="100%" border="0" cellspacing="5" cellpadding="5" class="no-style">

        <tr>
          <th>
            Ham
          </th>
        </tr>

        <tr>
          <td>
            <a href="http://www.dxsummit.fi/">Spots</a>
          </td>
        </tr>
		
        <tr>
          <td>
            <a href="http://www.reversebeacon.net/dxsd1.php?f=0&c=W2RDX&t=de">Reverse Beacon</a>
              <a href="http://reversebeacon.net/dxsd1/dxsd1.php?f=0&c=K2NNY&t=de" class="small">K2NNY</a>
          </td>
        </tr>
		
        <tr>
          <td>
            <a href="http://www.cordell.org/CI/CI_pages/CI_Overview.html">Cordell - Clipperton</a>
          </td>
        </tr>

        <tr>
          <td>
            <a href="http://www.websdr.org/">webSDR (Software Defined Radio)</a>
          </td>
        </tr>
		
        <tr>
          <td>
            <a href="https://lotw.arrl.org/lotwuser/default">LoTW</a>
          </td>
        </tr>
		
        <tr>
          <td>
            <a href="http://www.dxfc.org/">DXFC Feet Counties</a>
          </td>
        </tr>
		
        <tr>
          <td>
            <a href="http://www.latlong.net/">Latitude/Longitude Lookup</a>
          </td>
        </tr>
		
        <tr>
          <td>
            <a href="http://www.travellerspoint.com">Travellers Port</a>
          </td>
        </tr>
		
		
        <tr>
          <th>
            <br>Cam Servers
          </th>
        </tr>
        <tr>
          <td>
            <a href="http://www.stardot-tech.com/video_servers.html">Express 6 Video Servers</a>
              <a href="http://192.168.2.25" class="small">Server Room</a>
              <a href="http://192.168.2.95/index.html?cam1=on&cam2=on&size=1&columns=2&mode=0"
			    class="small">Dave's Office</a>
              <a href="http://192.168.1.101/index.html?cam1=on&cam2=on&cam3=on&cam4=on&cam5=on&cam6=on&size=1&columns=3&mode=0"
			    class="small">McConnellsville</a>
          </td>
        </tr>
        <tr>
          <td>
            Axis 2400+ Video Servers
              <a href="http://192.168.1.31" class="small">Axis-1</a>
              <a href="http://192.168.1.32" class="small">Axis-2</a>
              <a href="http://192.168.1.33" class="small">Axis-3</a>
              <a href="http://192.168.1.34" class="small">Axis-4</a>
          </td>
        </tr>
		
      </table>
    </td>
  </tr>
</table>
<br	/>
<br	/>

      <!--#include file="DynamicSections.inc"-->

</BODY>

</HTML>
