<!--#include virtual="includes/connection.asp"-->

<%

	ConnLocal.Execute "ha_Employee_logout_by_id(" & Session("ha_EmployeeID") & ")"

            Session("LoggedIn") = "False"
            Session("ha_fname") = ""
            Session("ha_lname") = ""
            Session("ha_shortname") = ""			
            Session("ha_login") = ""
            Session("ha_password") = ""
            Session("ha_EmpID") = ""
            Session("ha_EmployeeID") = ""
            Session("ha_EmpName") = ""
            Session("ha_bo_Authorization") = ""
			Session("ha_Permissions") = ""
            Session("eWeekEnding") = ""
                
    Response.Redirect strBasePath & "/includes/login.asp"
%>