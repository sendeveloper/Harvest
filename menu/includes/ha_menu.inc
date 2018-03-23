<%
   strPathMenu       =      strPathBase & "/menu/"
   
   if lcase(trim(ThisPage)) = "/includes/login.asp" then
%>

<!-- Login Menu -->
<div id="menu"></div>

<% elseif lcase(trim(ThisPage)) = "/timeclock/ha_timeclock.asp" or lcase(trim(ThisPage)) = "/payroll/ha_payrollsummary.asp" or lcase(trim(ThisPage)) = "/projects/ha_projectlist.asp" then %>

<!-- Back Only Menu -->
<div id="menu">
	<ul>
		<li style="float: right;"><a href="<%=Session("redirect")%>">Back</a></li>
	</ul>
</div>
<% else %>

<!-- Employee Menu -->
<div id="menu">
	<ul>
		<li><a href="<%=strPathMenu%>index.asp">Home</a>
          <ul>
            <li><a href="/menu/index.asp">Dashboard</a>
			<li><a href="http://crm.harvestamerican.info">CRM</a></li>
            <li><a href="http://www.number-it.com/Home/BackOffice/boAccounts.asp">Backoffice - Number-it</a></li>
            <li><a href="/ha_Backoffice/ha_Home.asp">Backoffice - Harvest</a></li>
            <li><a href="https://info.zip2tax.com/z2t_Backoffice/index.asp">Backoffice - Zip2tax</a></li>
            <li><a href="http://info2.zip2tax.com/z2t_Backoffice/index.asp">Backoffice - Zip2tax(Philly04)</a></li>
            <li style="display:none"><a href="http://legacy.zip2tax.com/Backoffice/z2t_login.asp">Backoffice - Zip2tax (Old)</a></li>
          </ul>
        </li>
		<li><a href="<%=strPathMenu%>sites.asp">Sites & Servers</a></li>
		<li><a href="<%=strPathMenu%>acct.asp">Accounting</a></li>
		<li><a href="<%=strPathMenu%>vndmrkt.asp">Vendors & Marketing</a></li>
		<li><a href="<%=strPathMenu%>webdev.asp">Web Dev</a></li>
		<li style="float: right;"><a href="#"><%=trim(Session("ha_fname"))%>&nbsp;<%=trim(Session("ha_lname"))%>&nbsp;&#9660</a>
			<ul class="last">
				<li><a href="/ha_BackOffice/Company/ha_timeclock<% If Session("ha_Permissions") = "99" Then Response.write "mgmt" Else Response.write "" %>.asp">Time Clock</a></li>
<% 
    If Session("ha_Permissions") > 59 Then
%>
	<li><a href="/ha_Backoffice/Company/ha_payroll.asp">Payroll</a></li>
	<% end if %>
				<li><a href="<%=strBasePath%>/includes/logout.asp">Logout</a></li>
			</ul>
		</li>
	</ul>
</div>

<% end if %>