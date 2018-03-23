<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
	<% if Session("ha_Permissions") < 60 then
		response.redirect("http://www.harvestamerican.info/menu/")
	end if %>
<%
     Session("redirect") = ""
     
     ThisPage = "https://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL")
%>

<!--#include virtual="includes/connection.asp"-->

<html xmlns="http://www.w3.org/1999/xhtml">

	<head>
		<title>Harvest American, Inc. | <% =Session("cMode") %> User</title>

		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
		<link href="<%=strPathMenuIncludes%>ha_menu.css" type="text/css" rel="stylesheet" />		
 
 </head>

	<body>
		
<!--#include virtual="menu/includes/ha_header.inc"-->

<!--#include virtual="menu/includes/ha_menu.inc"-->

		<div id="content">
		<form action="edituserpost.asp" method="get" name="frm">
			<div id="columnA">

			<div style="width:100%; margin:0 auto;">
									
				<div style="width:450; margin:0 auto;">
					<!--<h2><Edit><% =Session("cMode")%> User Info</h2>-->

				</div>

			</div>
			</div>

		<div id="columnB">
			<h2>Employees <a href="<%=strPathBase%>/settings/edituser.asp?cMode=Add" class="button" >+</a></h2>

		</div>
		</form>
		<div style="clear: both;">&nbsp;</div>

		<!--#include virtual="menu/includes/ha_footer.inc"-->
		
	</body>
</html>