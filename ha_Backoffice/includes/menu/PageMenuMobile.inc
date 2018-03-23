<tr align="center" valign="top">
  <td>
    <div id="moonmenu" class="halfmoon" style="padding-top: 1.25em;">
      <ul>
	    <li><a href="<%=strPathBase%>ha_Home.asp" 									rel="dropmenuHome">			<span>Home</span></a></li>
	    <li><a href="<%=strPathServers%>Summary/ha_Servers.asp" 					rel="dropmenuServers">		<span>Servers</span></a></li>
	    <li><a href="<%=strPathTelephones%>ha_TelephoneNumbers.asp" 				rel="dropmenuTelephones">	<span>Telephones</span></a></li>
	    <li><a href="<%=strPathNotifications%>ha_ServiceNotificationDetails.asp" 	rel="dropmenuNotifications"><span>Notifications</span></a></li>
        <li><a href="<%=strPathEmployees%>Employees/ha_Employees.asp" 				rel="dropmenuEmployees">	<span>Employees</span></a></li>
        <li><a href="<%=strPathCompany%>Accomplishments/ha_Accomplishments.asp" 	rel="dropmenuCompany">		<span>Company</span></a></li>
      </ul>
    </div>    
    
    <!--0th drop down menu -->                                                   
    <div id="dropmenuHome" class="dropmenudiv_e" style="width: 15em;">
      <a href="/menu/index.asp">Dashboard</a>
      <a href="<%=strPathBase%>ha_Home.asp">Maintenance</a>
      <a href="https://www.zip2tax.info/z2t_Backoffice/index.asp">Zip2tax</a>
      <a href="http://legacy.zip2tax.com/Backoffice/z2t_login.asp">Old Zip2tax</a>
      <a href="http://www.number-it.com/Home/BackOffice/boAccounts.asp">Backoffice</a>
    </div>

    <!--1st drop down menu -->                                                   
    <div id="dropmenuServers" class="dropmenudiv_e" style="width: 18em;">
      <a href="<%=strPathServers%>Summary/ha_Servers.asp">Summary</a>
      <a href="<%=strPathServers%>ha_Servers_Archives.asp">Archives</a>
      <a href="<%=strPathServers%>Backups/ha_Servers_Backups.asp">Backups</a>
      <a href="<%=strPathServers%>Backups/ha_Servers_Backups_Database.asp">Backups (Databases)</a>
      <a href="<%=strPathServers%>Domains/ha_Domains.asp">Domains</a>
      <a href="<%=strPathServers%>ha_WebpageHits.asp">Webpage Hits</a>
      <a href="<%=strPathServers%>TableDistribution/ha_TableDistribution.asp">Table Distribution</a>
      <a href="<%=strPathServers%>ha_ServerDiskSpace.asp">Server Disk Space</a>
      <a href="<%=strPathServers%>ha_Servers_RMC.asp">Remote Management (Reboot)</a>
    </div>
    //Decided not to use
    <!--2nd drop down menu -->                                                   
    <div id="dropmenuTelephones" class="dropmenudiv_e" style="width: 15em;">
      <a href="<%=strPathTelephones%>ha_TelephoneNumbers.asp">Numbers</a>
      <a href="<%=strPathTelephones%>ha_TelephoneCallLog.asp">Call Log</a>
      <a href="<%=strPathTelephones%>ha_TelephoneStatistics.asp">Statistics</a>
      <a href="<%=strPathTelephones%>ha_TelephoneVPN.asp">VPN Connections</a>
    </div>
    
    <!--3rd drop down menu -->                                                   
    <div id="dropmenuNotifications" class="dropmenudiv_e" style="width: 15em;">
      <a href="<%=strPathNotifications%>ha_ServiceNotificationStatus.asp">Status</a>
      <a href="<%=strPathNotifications%>ha_ServiceNotificationEdit.asp">Edit</a>
      <a href="<%=strPathNotifications%>ha_ServiceNotificationDetails.asp">Details</a>
    </div>
    
    <!--4th drop down menu-->
    <div id="dropmenuEmployees" class="dropmenudiv_e" style="width: 15em;">
	  <a href="<%=strPathEmployees%>Employees/ha_Employees.asp">Employees</a>
	  <a href="<%=strPathEmployees%>ha_InternTimeline.asp">Intern Timeline</a>
      <a href="<%=strPathEmployees%>ha_employees_add.asp">Add/Edit</a>
      <a href="<%=strPathCompany%>ha_timeclock<%=iif(Session("ha_Permissions") = "99", "mgmt","")%>.asp">Time Clock</a>
      <%=ListIfPermitted(11, "<a href='" & strPathEmployees & "ha_PayrollHours.asp'>Payroll Hours</a>")%>
      <a href="<%=strPathEmployees%>Calendar/ha_Calendar.asp">Vacation Calendar</a>
	  <a href="<%=strPathEmployees%>ha_EmployeeLoginStatus.asp">Login Status</a>
      <%=ListIfPermitted(10, "<a href='" & strPathEmployees & "ha_EmployeeSetup.asp'>Setup Procedures</a>")%>
      <a href="<%=strPathEmployees%>PermissionDefinitions/ha_PermissionDefinitions.asp">Permission Definitions</a>
    </div>
    
    <!--5th drop down menu-->
    <div id="dropmenuCompany" class="dropmenudiv_e" style="width: 15em;">
      <a href="<%=strPathCompany%>Accomplishments/ha_Accomplishments.asp">Accomplishments</a>
      <a href="<%=strPathCompany%>DocumentMaintenance/ha_DocumentMaintenance.asp">Document Maintenance</a>
      <a href="<%=strPathCompany%>Products/ha_Products.asp">Products</a>
      <a href="<%=strPathCompany%>PasswordManagement/ha_PasswordManagement.asp">Password Management</a>
      <a href="<%=strPathCompany%>ToDo/ha_ToDo.asp">To Do List</a>
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
