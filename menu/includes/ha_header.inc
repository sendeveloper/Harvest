<%
   if lcase(trim(ThisPage)) = "/includes/login.asp" then
%>

<!-- Login Header -->
<div id="header">
	<h1>Harvest American, Inc.</h1>
	<h2>Login</h2>
</div>

<%
   elseif lcase(trim(ThisPage)) = "/timeclock/ha_timeclock.asp" then
%>

<!-- Time Clock Header -->
<div id="header">
	<h1>Harvest American, Inc.</h1>
	<h2>Employee Time Clock</h2>
</div>

<%
   elseif lcase(trim(ThisPage)) = "/payroll/ha_payrollsummary.asp" then
%>

<!-- Time Clock Header -->
<div id="header">
	<h1>Harvest American, Inc.</h1>
	<h2>Payroll Administration</h2>
</div>

<%
   else
%>

<!-- Generic Header -->
<div id="header">
	<h1>Harvest American, Inc.</h1>
	<h2>Dashboard</h2>
</div>

<%
   end if
%>