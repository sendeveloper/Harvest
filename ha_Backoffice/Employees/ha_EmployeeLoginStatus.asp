<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include virtual="includes/connection.asp"-->
<!--#include virtual="ha_BackOffice/includes/Config.asp"-->
<%
	ColorTab = 4
	PageHeading = "Employee Login Status"
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
  <link href="calendar.css" rel="stylesheet" type="text/css" />
  <script type="text/javascript" src="calendar_date_picker.js"></script>
  <title>Harvest American Backoffice - Employee Login Status</title>
  <!--<meta http-equiv="REFRESH" content="0; url==strPathEmployees%>ha_employees_add.asp" />-->
  <link href="<%=strPathIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet" />
  <link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>menu/dropdowntabfiles/halfmoontabs.css" />
  <script type="text/javascript" src="<%=strPathIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>
</head>

<body class="gray_desktop">
  <table cellpadding="0" cellspacing="0" class="MainBody" style="margin: 20px auto 0;width: 1200px !important;">
    <tr>
      <td>
        <!--#include virtual="ha_Backoffice/includes/bodyparts/ha_PageHeading.inc"-->
      </td>
    </tr>
	
	<tr><td>&nbsp;</td></tr>
	<tr><td>&nbsp;</td></tr>
	
    <tr>
      <td align="center">
        <table width="800" cellpadding="2" cellspacing="2">
		  <tr>
		    <td>
			  This is a list of Session Variables that were created upon login that is information
			  carried by the backoffice about you.  This page is usefull for programing and diagnosing
			  problems.<br><br>
			</tr>
		  </tr>
		  
		  <tr>
		    <td style="font-size: 11pt;">
			  <span style="font-weight: bold; color: #336600;">Session Variables</span><br>
			  <hr>
			</td>
		  </tr>
		  
		  <tr>
			<td>
			  <table width="100%" cellpadding="1" cellspacing="1">
			    <tr><td>Session("LoggedIn")				=</td><td><b><%=Session("LoggedIn")%>				</b></td></tr>
				<tr><td>Session("ha_fname") 			=</td><td><b><%=Session("ha_fname")%>				</b></td></tr>
				<tr><td>Session("ha_lname") 			=</td><td><b><%=Session("ha_lname")%>				</b></td></tr>
				<tr><td>Session("ha_ShortName")			=</td><td><b><%=Session("ha_ShortName")%>			</b></td></tr>
				<tr><td>Session("ha_login") 			=</td><td><b><%=Session("ha_login")%>				</b></td></tr>
				<tr><td>Session("ha_password") 			=</td><td><b><%=Session("ha_password")%>			</b></td></tr>
				<tr><td>Session("ha_EmployeeID") 		=</td><td><b><%=Session("ha_EmployeeID")%>			</b></td></tr>
				<tr><td>Session("ha_EmpID") 			=</td><td><b><%=Session("ha_EmpID")%>				</b></td>
					<td><span style="color: #C0C0C0;">(trying to depreciate, use ha_EmployeeID)</span>		</b></td></tr>
				<tr><td>Session("ha_EmpName") 			=</td><td><b><%=Session("ha_EmpName")%>				</b></td></tr>
				<tr><td>Session("ha_bo_Authorization")	=</td><td><b><%=Session("ha_bo_Authorization")%>	</b></td></tr>
				<tr><td>Session("ha_Permissions") 		=</td><td><b><%=Session("ha_Permissions")%>			</b></td></tr>
				
				<tr><td>&nbsp;</td></tr>
				<tr><td style="font-weight: bold; color: #336600;">New Zip2Tax Backoffice</td></tr>
				<tr><td>Session("z2t_login") 			=</td><td><b><%=Session("z2t_login")%>				</b></td></tr>
				<tr><td>Session("z2t_loggedin") 		=</td><td><b><%=Session("z2t_loggedin")%>			</b></td></tr>
				<tr><td>Session("z2t_status") 			=</td><td><b><%=Session("z2t_status")%>				</b></td></tr>
				<tr><td>Session("z2t_AccountID") 		=</td><td><b><%=Session("z2t_AccountID")%>			</b></td></tr>
				<tr><td>Session("z2t_UserName") 		=</td><td><b><%=Session("z2t_UserName")%>			</b></td></tr>
				<tr><td>Session("z2t_ShortName") 		=</td><td><b><%=Session("z2t_ShortName")%>			</b></td></tr>

				<tr><td>&nbsp;</td></tr>

				<tr><td>Session("Login") 				=</td><td><b><%=Session("Login")%>					</b></td></tr>
				<tr><td>Session("Authority") 			=</td><td><b><%=Session("Authority")%>				</b></td></tr>
				<tr><td>Session("emp_fname") 			=</td><td><b><%=Session("emp_fname")%>				</b></td></tr>
				<tr><td>Session("Status") 				=</td><td><b><%=Session("Status")%>					</b></td></tr>
				<tr><td>Session("lAccountID") 			=</td><td><b><%=Session("lAccountID")%>				</b></td></tr>
			  </table>
			</td>
		  </tr>
		
		  <tr>
		    <td>
			  <hr>
			</td>
		  </tr>
		  
		</table>
      </td>
    </tr>
	
	<tr><td>&nbsp;</td></tr>
	<tr><td>&nbsp;</td></tr>

  </table>
  <!--#include virtual="ha_BackOffice/includes/BodyParts/ha_PageTail.inc"-->
</body>
</html>

