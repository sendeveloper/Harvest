<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<!--#include virtual="includes/connection.asp"-->
<%
	DIM rs
	DIM SQL
			
	set rs = server.createObject("ADODB.Recordset")
	SQL="ha_DaveMenu" 
	rs.open SQL,objcon, 3, 3, 4

	NumberOfSeconds= rs("NumberOfSeconds")
	StatesLeft = rs("StatesLeft")
	Left2Split = rs("Left2Split")

	rs.close
%>
<html>
<head>
    <title>Dave's Menu</title>

    <link href="<%=strPathIncludes%>HarvestAmerican.css" type="text/css" rel="stylesheet">
    <script language="JavaScript" src="<%=strPathIncludes%>haFunctions.js"></script>

    <style type="text/css">

	td
		{
		font-family: Arial, Helvetica, sans-serif; 
		font-size: 12px; 
		font-style: normal; 
		line-height: normal; 
		font-weight: normal; 
		font-variant: normal; 
		text-transform: none;
		color: #000000;
		text-decoration: none
		}

	th
		{
		font-size: 14px;
		}

	a
		{
		font-size: 12px;
		}

	a.small
		{
		font-size: 9px;
		padding-left: 5px;
		}

    </style>

</head>

<body>

<table width='100%' border='2' cellspacing='5' cellpadding='5'>
  <tr width='100%'>
    <td colspan='4' align='center'>
      <font Size='4'><b>Dave's Menu</b></font>
    </td>
  </tr>

  <tr valign='top'>

    <td width='25%' align='center'>
      <table width='100%' border='0' cellspacing='5' cellpadding='5'>
        <tr>
          <th>
            Orders
          </th>
        </tr>
        <tr>
          <td>
            <a href='https://ebc.cybersource.com/ebc/login/Login.do'>CyberSource</a>
          </td>
        </tr>
        <tr>
          <td>
            <a href='http://www.snscart.com/login.htm'>Shopping Cart</a>
          </td>
        </tr>
        <tr>
          <td>
            <a href='http://www.number-it.com'>Number-it</a>
            <a href='http://www.number-it.com/Home/BackOffice/boLogin.asp?lname=<%=Session("ha_login")%>&pass=<%=Session("ha_password")%>'
              class="small">Backoffice</a>
          </td>
        </tr>
        <tr>
          <td>
            <a href='http://www.amazon.com/gp/seller-account/management/your-account.html/002-3179982-7492809'>
              Amazon Seller Account</a>
          </td>
        </tr>


        <tr>
          <th>
            <br>Shipping
          </th>
        </tr>
        <tr>
          <td>
            <a href='https://sss-web.usps.com/cns/landing.do'>UPS</a>
          </td>
        </tr>
        <tr>
          <td>
            <a href='http://www.usps.com'>USPS</a>
          </td>
        </tr>

        <tr>
          <th>
            <br>Your Web Vendors
          </th>
        </tr>
        <tr>
          <td>
            <a href='https://adwords.google.com/select/Login?sourceid=awo&subid=ww-en-et-ads1_e&hl=en_us'>
              Google AdWords</a>
            <a href='https://www.google.com/webmasters/tools/siteoverview?hl=en&msg=2&pli=1' class="small">
              Sitemaps</a>
          </td>
        </tr>
        <tr>
          <td>
            <a href='https://adcenter.microsoft.com'>
              MSN</a>
          </td>
        </tr>
        <tr>
          <td>
            <a href='https://login22.marketingsolutions.yahoo.com'>
              Yahoo</a>
          </td>
        </tr>
        <tr>
          <td>
            <a href='http://ConstantContact.com'>Constant Contact</a>
            <a href='Content/Constant Contact Email Procedure.rtf' class="small">How</a>
          </td>
        </tr>

      </table>
    </td>

    <td width='25%' align='center'>
      <table width='100%' border='0' cellspacing='5' cellpadding='5'>

        <tr>
          <th>
            Banking
          </th>
        </tr>
        <tr>
          <td>
            <a href='https://www.accessfcu.org/onlineserv/HB/Signon.cgi'>Access Federal CU</a>
          </td>
        </tr>
        <tr>
          <td>
            <a href='https://sitekey.bankofamerica.com/sas/signonSetup.do'>Bank of America</a>
          </td>
        </tr>
        <tr>
          <td>
            <a href='http://www.nbtbank.com'>NBT</a>
          </td>
        </tr>

        <tr>
          <th>
            <br>Bookkeeping
          </th>
        </tr>
        <tr>
          <td>
            <a href="<%=strPathBase%>/timeclock/ha_timeclock.asp">
              Employee Hours</a>
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
            <a href='https://wwws.ameritrade.com/apps/LogIn'>
              TD Ameritrade</a>
          </td>
        </tr>
        <tr>
          <td>
            <a href='https://www.fidelity.com/'>
              Fidelity</a>
          </td>
        </tr>
        <tr>
          <td>
            <a href='http://finance.yahoo.com/q/cp?s=%5EDJI'>
              Dow Components</a>
          </td>
        </tr>

      </table>
    </td>



    <td width='25%' align='center'>
      <table width='100%' border='0' cellspacing='5' cellpadding='5'>
        <tr>
          <th>
            Your Websites
          </th>
        </tr>

        <tr>
          <td>
            <a href='http://www.CamdenBudgetStorage.com/home/login/logpost.asp?lname=dave&pass=go4it'>
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
            <a href='http://www.EyeOnCamden.com'>Eye On Camden</a>
          </td>
        </tr>

        <tr>
          <td>
            <a href='http://www.HarvestAmerican.com'>Harvest American</a>
              <a href='http://www.HarvestAmerican.info/ControlPanel/stats/ha_Stats.asp' class="small">Stats</a>
              <a href='http://www.harvestamerican.com/wholesale/tickets/'class="small">Wholesale</a>
              <a href='http://helpdesk.harvestamerican.net' class="small">Help Desk</a>
         </td>
        </tr>

        <tr>
          <td>
            <a href='http://www.number-it.com'>Number-it</a>
            <a href='http://www.number-it.com/Home/BackOffice/boLogin.asp?lname=<%=Session("ha_login")%>&pass=<%=Session("ha_password")%>'
              class="small">Backoffice</a>
          </td>
        </tr>

        <tr>
          <td>
            <a href='http://www.NYSThruway.com'>NYSThruway.com</a>
          </td>
        </tr>

        <tr>
          <td>
            <a href="http://www.WJ2O.com?LoggedIn=True">WJ2O.com</a>
              <a href='Content/How To Add To WJ2O Database.rtf' class="small">Update Process</a>
          </td>
        </tr>

        <tr>
          <td>
            <a href='http://www.zip2tax.com'>Zip2Tax</a>
              <a href="http://www.zip2tax.com/backoffice/z2t_Login.asp?lname=davewj2o&pass=get2it" class="small">
                Backoffice</a>
              <a href="http://www.harvestamerican.us/common/reference/z2t_dotnums.txt" class="small">
                State Numbers</a>
              <a href='Content/z2t Monthly Updates Procedure.rtf' class="small">
                Update Process</a>
              <a href='Content/z2t Month End Procedures.rtf' class="small">
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
            <a href='<%=strPathBase%>/ControlPanel/cp_Servers.asp'>Harvest Control Panel</a>
              <a href='<%=strPathBase%>/ControlPanel/cp_Servers.asp?lname=nopassword' class="small">No Password</a>
          </td>
        </tr>
        <tr>
          <td>
            <a href='https://68.178.202.54:8443/login.php3'>Plesk on Harvest</a>
          </td>
        </tr>
        <tr>
          <td>
            <a href='https://orbit.theplanet.com/Network/OrbitGraphDetail.aspx?id=155931'>SoftLayer (The Planet) Control Panel</a>
          </td>
        </tr>
        <tr>
          <td>
            <a href='https://unity.serverintellect.com'>Server Intellect Control Panel</a>
          </td>
        </tr>
        <tr>
          <td>
            <a href='http://www.ultradns.net'>Neustar (DNS Server)</a>
          </td>
        </tr>


      </table>
    </td>

    <td width='25%' align='center'>
      <table width='100%' border='0' cellspacing='5' cellpadding='5'>

        <tr>
          <th>
            Clients
          </th>
        </tr>

        <tr>
          <td>
            <a href='http://www.tlls.org'>Leukemia</a>
              <a href='Contacts_leuk.asp' class="small">Contacts</a>
              <a href='https://webmail.lls.org' class="small">Mail</a>
              <a href='http://SupportCenter.lls.org' class="small">SupportCenter</a>
              <a href='https://vpn.lls.org/webvpn.html' class="small">VPN</a>
              <a href='https://secure.lls.org/Citrix/MetaFrame/default/default.aspx' class="small">Citrix</a>
              <a href='http://jira.tlls.net/' class="small">JIRA</a>
          </td>
        </tr>

        <tr>
          <td>
            <a href='http://www.gotomypc.com'>ccfa Production Server</a>
              <a href='Content/How To Login To ccfa Prod Server.txt' class="small">How to log on to CCFA server</a>
              <a href='Content/Jackie_GoToMyPC_connection.txt' class="small">How to log on to Jackie's PC</a>
              <a href='Contacts_ccfa.asp' class="small">Contacts</a>
          </td>
        </tr>

        <tr>
          <td>
            <a href='https://www.filecollegeinfo.com'>cts Website</a>
              <a href='http://test.filecollegeinfo.com' class="small">Test</a>
              <a href='Contacts_cts.asp' class="small">Contacts</a>
          </td>
        </tr>

        <tr>
          <td>
            <a href='https://www.Mkombozi.net'>Mkombozi</a>
          </td>
        </tr>

        <tr>
          <th>
            <br>Reference
          </th>
        </tr>
        <tr>
          <td>
            <a href='Content/TelephoneWiringDiagrams.xls'>Telephone Wiring Diagrams</a>
          </td>
        </tr>
        <tr>
          <td>
            <a href='http://www.htmlhelp.com/reference/css/properties.html'>WDG CSS Properties</a>
          </td>
        </tr>
      </table>
    </td>

  </tr>
</table>
<table width='100%' border='0' cellspacing='5' cellpadding='5' align="left">
  <tr>
    <td width="25%" valign="top">
<% if int(NumberOfSeconds) > 120 then %>
    <span style="color: red; font-weight: bold"> WARNING! "ha_Requirements_process" IS NOT EXECUTING!</span><br><hr>
<% else %>
    ha_Requirements_process ran <%=NumberOfSeconds%> seconds ago<br><hr>
<% end if %>
    <%=StatesLeft%> 
      <% if StatesLeft = "1" then%>
        State left to update this month<br><hr>
      <% else %>
        States left to update this month<br><hr>
      <% end if %>	
    </td>

    <td width="25%">&nbsp;
      
    </td>

    <td width="50%">
      Session("LoggedIn") = <%=Session("LoggedIn")%><br>
      Session("ha_fname") = <%=Session("ha_fname")%><br>
      Session("ha_lname") = <%=Session("ha_lname")%><br>
      Session("ha_login") = <%=Session("ha_login")%><br>
      Session("ha_password") = <%=Session("ha_password")%><br>
      Session("ha_EmpID") = <%=Session("ha_EmpID")%><br>
      Session("ha_bo_Authorization") = <%=Session("ha_bo_Authorization")%><br>
    </td>
  </tr>

</table>

</BODY>

</HTML>