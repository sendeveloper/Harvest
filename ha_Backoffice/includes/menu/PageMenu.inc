<tr align="center" valign="top">
  <td>
    <div id="moonmenu" class="halfmoon" style="padding-top: 1.25em;">
      <ul>
	    <li><a href="/ha_Backoffice/ha_Home.asp" 									rel="dropmenuHome">			<span>Home</span></a></li>
	    <li><a href="<%=strPathServers%>Summary/ha_Servers.asp" 					rel="dropmenuServers">		<span>Servers</span></a></li>
	    <li><a href="<%=strPathTelephones%>ha_TelephoneNumbers.asp" 				rel="dropmenuTelephones">	<span>Telephones</span></a></li>
	    <li><a href="/ha_Backoffice/Operations/z2t_Orders/z2t_Operations_orders.asp"	rel="dropmenuOperations">	<span>Operations</span></a></li>
        <li><a href="<%=strPathEmployees%>Employees/ha_Employees.asp" 				rel="dropmenuEmployees">	<span>Employees</span></a></li>
        <li><a href="<%=strPathCompany%>Accomplishments/ha_Accomplishments.asp" 	rel="dropmenuCompany">		<span>Company</span></a></li>
        <li><a href="/ha_Backoffice/Statistics/ha_Statistics_Activity.asp"		 	rel="dropmenuStatistics">	<span>Statistics</span></a></li>
		<li><a href="/ha_Backoffice/Utilities/ha_ConstantContact.asp"		 		rel="dropmenuUtilities">	<span>Utilities</span></a></li>
	    <li><a href="/ha_Backoffice/Network/ha_Network.asp" 						rel="dropmenuNetwork">		<span>Network</span></a></li>
	    <li><a href="/ha_Backoffice/Notifications/ha_Notifications.asp" 			rel="dropmenuNotifications">	<span>Notifications</span></a></li>
	    <li><a href="/ha_Backoffice/Websites/WordPress/ha_WordPress.asp" 			rel="dropmenuWebsites">		<span>Websites</span></a></li>
      </ul>
    </div>

    <!--0th drop down menu -->                                                   
    <div id="dropmenuHome" class="dropmenudiv_e" style="width: 15em;">
      <a href="http://www.harvestamerican.info/menu/index.asp">Dashboard</a>
	  <a href="http://crm.harvestamerican.info">CRM (Customer Relationship Management)</a>
      <a href="http://www.number-it.com/Home/BackOffice/boAccounts.asp">Backoffice - Number-it</a>
      <a href="<%=strPathBase%>ha_Home.asp">Backoffice - Harvest</a>
      <a href="http://info.zip2tax.com/z2t_Backoffice/home/z2t_home.asp">Backoffice - Zip2Tax</a>
      <a href="http://legacy.zip2tax.com/Backoffice/z2t_login.asp" style="display:none;">Backoffice -  Zip2Tax (Old)</a>
    </div>	

    <!--1st drop down menu -->                                                   
    <div id="dropmenuServers" class="dropmenudiv_e" style="width: 18em;">
      <a href="<%=strPathServers%>Summary/ha_Servers.asp">Summary</a>
      <a href="<%=strPathServers%>ServerMaintenance/ha_ServerMaintenance.asp">Server Maintenance Log</a>
      <a href="<%=strPathServers%>ha_Servers_Archives.asp">Archives</a>
      <a href="<%=strPathServers%>Backups/ha_Servers_Backups.asp">Backups</a>
      <a href="<%=strPathServers%>Backups/ha_Servers_Backups_Database.asp">Backups (Databases)</a>
      <a href="<%=strPathServers%>ha_WebpageHits.asp">Webpage Hits</a>
	  <a href="<%=strPathServers%>ha_ShoppingCarts.asp">Shopping Cart Process</a> 
      <a href="<%=strPathServers%>TableDistribution/ha_TableDistribution.asp">Table Distribution</a>
      <a href="<%=strPathServers%>TableDistribution/ha_TableDistribution_Q_Process.asp">Table Distribution Q Process</a>
      <a href="<%=strPathServers%>ha_TableLocations.asp">Table Locations</a>
      <a href="<%=strPathServers%>ha_ServerDiskSpace.asp">Server Disk Space</a>
      <a href="<%=strPathServers%>ha_Servers_RMC.asp">Remote Management (Reboot)</a>
      <a href="<%=strPathServers%>ha_LoadBalancer.asp">Load Balancer</a>
      <a href="/ha_Backoffice/Servers/Domains/ha_Domains.asp">Domains</a>
      <a href="/ha_Backoffice/Servers/ha_TableDownloads.asp">Table Downloads</a>
      <a href="/ha_Backoffice/Servers/VideoServers/ha_VideoServers.asp">Video Servers</a>
      <a href="<%=strPathServers%>ha_Websites.asp">Websites</a>
      <a href="/ha_Backoffice/Servers/DatabaseClasses/ha_DatabaseClasses.asp">Database Classes</a>
    </div>
    
    <!--2nd drop down menu -->                                                   
    <div id="dropmenuTelephones" class="dropmenudiv_e" style="width: 15em;">
      <a href="<%=strPathTelephones%>ha_TelephoneNumbers.asp">Numbers</a>
      <a href="<%=strPathTelephones%>ha_TelephoneCallLog.asp">Call Log</a>
      <a href="<%=strPathTelephones%>ha_TelephoneStatistics.asp">Statistics</a>
	  <a href="<%=strPathTelephones%>ha_LoggingProcess.asp">Logging Process</a>
      <a href="<%=strPathTelephones%>ha_TelephoneVPN.asp">VPN Connections</a>
      <a href="/ha_Backoffice/Telephones/PBX/ha_PBX_Information.asp">PBX Information</a>
    </div>
    
    <!--3rd drop down menu -->                                                   
    <div id="dropmenuOperations" class="dropmenudiv_e" style="width: 18em;">
        <a href="/ha_Backoffice/Operations/z2t_Orders/z2t_Operations_Orders.asp">Zip2Tax Order Process</a>
    </div>
    
    <!--4th drop down menu-->
    <div id="dropmenuEmployees" class="dropmenudiv_e" style="width: 15em;">
	  <a href="<%=strPathEmployees%>Employees/ha_Employees.asp">Employees</a>
	  <a href="<%=strPathEmployees%>Calendar/ha_EmployeeCalendar.asp">Calendar</a>
	  <a href="<%=strPathEmployees%>Email/Email/ha_EmployeeEmail.asp">E-mail</a>
	  <a href="<%=strPathEmployees%>Email/Groups/ha_EmployeeEmailGroups.asp">E-mail Groups</a>
	  <a href="<%=strPathEmployees%>Email/Aliases/ha_EmployeeEmailAliases.asp">E-mail Aliases</a>
	  <a href="<%=strPathEmployees%>Employees/ha_EmployeeForms.asp">Employee Forms</a>
	  <a href="<%=strPathEmployees%>ha_InternTimeline.asp">Intern Timeline</a>
      <a href="<%=strPathCompany%>ha_timeclock<%=iif(Session("ha_Permissions") = "99", "mgmt","")%>.asp">Time Clock</a>
      <%=ListIfPermitted(11, "<a href='" & strPathEmployees & "ha_PayrollHours.asp'>Payroll Hours</a>")%>
	  <a href="<%=strPathEmployees%>ha_EmployeeLoginStatus.asp">Login Status</a>
      <%=ListIfPermitted(10, "<a href='" & strPathEmployees & "ha_EmployeeAddRemove.asp'>Add/Remove Credentials</a>")%>
      <a href="<%=strPathEmployees%>PermissionDefinitions/ha_PermissionDefinitions.asp">Permission Definitions</a>
      <a href="<%=strPathEmployees%>JobDescriptions/ha_JobDescriptionTitles.asp">Job Description Titles</a>
      <a href="<%=strPathEmployees%>JobDescriptions/ha_JobDescriptionTasks.asp">Job Description Tasks</a>
    </div>
    
    <!--5th drop down menu-->
    <div id="dropmenuCompany" class="dropmenudiv_e" style="width: 15em;">
      <a href="<%=strPathCompany%>Accomplishments/ha_Accomplishments.asp">Accomplishments</a>
      <a href="<%=strPathCompany%>DocumentMaintenance/ha_DocumentMaintenance.asp">Document Maintenance</a>
      <a href="<%=strPathCompany%>Notes/ha_Company_Notes.asp">Company Notes</a>
      <a href="<%=strPathCompany%>Products/ha_Products.asp">Products</a>
      <a href="<%=strPathCompany%>PasswordManagement/ha_PasswordManagement.asp">Password Management</a>
      <a href="<%=strPathCompany%>ToDo/ha_ToDo.asp">To Do List</a>
    </div>
	
    <!--6th drop down menu-->
    <div id="dropmenuStatistics" class="dropmenudiv_e" style="width: 15em;">
      <a href="/ha_Backoffice/Statistics/ha_Statistics_Activity.asp">Activity</a>
    </div>
	
	<!--7th drop down menu-->
    <div id="dropmenuUtilities" class="dropmenudiv_e" style="width: 15em;">
      <a href="/ha_Backoffice/Utilities/ha_ConstantContact.asp">Constant Contact Export</a>
	  <a href="/ha_Backoffice/Utilities/ha_StateLinks.asp">Raffle Ticket State Links</a>
    </div>


    <!--8th drop down menu -->                                                   
    <div id="dropmenuNetwork" class="dropmenudiv_e" style="width: 18em;">
        <a href="/ha_Backoffice/Network/ip/ha_staticIP.asp">Static IP Mappings</a>
        <a href="/ha_Backoffice/Network/ip/ha_dns.asp">DNS Mappings</a>
        <a href="/ha_Backoffice/Network/ActiveDirectory/ha_ActiveDirectory.asp">Active Directory</a>
        <a href="/ha_Backoffice/Network/certificates/ha_Certificates.asp">SSL Certificates</a>
		<a href="/ha_Backoffice/Network/backup/ha_Backup.asp">Backups</a>
    </div>
    
    <!--9th drop down menu -->                                                   
    <div id="dropmenuNotifications" class="dropmenudiv_e" style="width: 18em;">
        <a href="/ha_Backoffice/Notifications/ha_Notifications.asp">Notifications</a>
        <a href="/ha_Backoffice/Notifications/ha_NotificationAlerts.asp">Notification Alerts</a>
        <a href="/ha_Backoffice/Notifications/ha_NotificationSensors.asp">Notification Sensors</a>
    </div>
	
    <!--10th drop down menu -->                                                   
    <div id="dropmenuWebsites" class="dropmenudiv_e" style="width: 18em;">
        <a href="/ha_Backoffice/Websites/OnlinePaymentProcess/ha_OnlinePaymentProcess.asp">Online Payment Process</a>
        <a href="/ha_Backoffice/Websites/OrderUploadProcess/ha_OrderUploadProcess.asp">Order Upload Process</a>
		<a href="/ha_Backoffice/Websites/PaymentProcessorInterface/ha_PaymentProcessorinterface.asp">Payment Processor Interface</a>
        <a href="/ha_Backoffice/Websites/TrialSubscriptionsProcess/ha_TrialSubscriptionsProcess.asp">Trial Subscriptions Process</a>
		<a href="/ha_Backoffice/Websites/WebsiteFlowchart/ha_WebsiteFlowchart.asp">Website Flowchart</a>
        <a href="/ha_Backoffice/Websites/WordPress/ha_WordPress.asp">WordPress</a>
    </div>
	
    <script type="text/javascript">
      tabdropdown.init("moonmenu", <%=ColorTab%>)
    </script>
    
  </td>
</tr>

<%  
	Function ListIfPermitted (pID, strURL)
		If HasPermission("Read", pID) Then
			Response.Write strURL
		End If
	End Function
%>
