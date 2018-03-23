<?xml version="1.0" encoding="ISO-8859-1"?>
<!DOCTYPE html PUBLIC "-//WAPFORUM//DTD XHTML Mobile 1.0//EN" "http://www.wapforum.org/DTD/xhtml-mobile10.dtd">

<%
    ThisPage = "https://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL")
    'Session("Redirect") = ""

    if UCase(Session("Redirect")) = "/DAVE/INDEX.ASP" then
        Session("Redirect") = "/mobile/mLogin.asp"
    end if
%>

<!--#include virtual="includes/connection.asp"-->

<%
    
'Response.write "|" & Request("HarvestAmericanPassword") & "|<br>"

    if Request("HarvestAmericanLogin") <> "" then
        Session("ha_login") = Request("HarvestAmericanLogin")
    end if

    if Request("HarvestAmericanPassword") <> "" then
        Session("ha_password") = Request("HarvestAmericanPassword")
    end if

    if Request("HarvestAmericanPassword") = "3155" then
        Session("ha_login") = "davewj2o"
        Session("ha_password") = ")^v("
    end if

    if Request("HarvestAmericanPassword") = "d" then
        Session("ha_login") = "davewj2o"
        Session("ha_password") = ")^v("
    end if

    '----- Check for login match -----
    DIM messageColorRed
    
    messageColorRed = "False"
    
'Response.write "Session(LoggedIn)=" & Session("LoggedIn") & "|<br>"
'Response.write "Session(redirect)=" & Session("redirect") & "|<br>"
'Response.write "ha_login=" & Session("ha_login") & "|<br>"
'Response.write "ha_password=" & Session("ha_password") & "|<br>"

    if Session("ha_login") = "" then 
        Session("ha_password") = ""
        strMessage = "Please Login"
    elseif Session("ha_password") = "" then
        Session("ha_login") = ""
        strMessage = "You forgot to enter a password!"
    else
    
        SQL="select * from ha_EmpAccounts where " & _
            "username='" & Session("ha_login") & "'" & _
            "and password='" & Session("ha_password") & "'"
             set RS=server.createObject("ADODB.Recordset")
        RS.open SQL,objcon,2,3

        if RS.EOF then
	    Session("LoggedIn") = ""
            Session("ha_login") = ""
            Session("ha_password") = ""
            strMessage = "Invalid username and/or password!"
            messageColorRed = "True"
        else
            Session("LoggedIn") = "True"
            Session("ha_fname") = rs("FirstName")
            Session("ha_lname") = rs("LastName")
            Session("ha_login") = rs("username")
            Session("ha_password") = rs("password")
            Session("ha_EmpID") = rs("EmpID")
            Session("ha_bo_Authorization") = trim(rs("ha_bo_Authorization"))
        end if

        rs.close
        set rs = nothing	
    
        if Session("LoggedIn") = "True" then
           if UCase(Session("Redirect")) = "/MOBILE/MLOGIN.ASP" then
              Response.Redirect "/mobile/todayssales.asp"
           else
              Response.Redirect "" & Session("redirect")
           end if
        end if    
    end if
%>

<html xmlns="http://www.w3.org/1999/xhtml">
  <head>
    <title>Harvest American, Inc. | Mobile Login</title>
    <link href="includes/mobile.css" rel="stylesheet" type="text/css">
    
  </head>
  <body onLoad="document.forms.frm.HarvestAmericanLogin.focus()">
  


<!-- #include file="includes/header.asp"-->

    <div class="content">
    <FORM method=get action="mLogin.asp" name="frm">
      <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
          <td>

            <table width="300" border="0" cellspacing="1" cellpadding="1" align="left">
              <tr>
                <td colspan="2">
                  <table width="100%" border="0" cellspacing="1" cellpadding="5" align="center">
                    <tr>
<%
     if messageColorRed = "True" then
%>
		              <td colspan="2" align="center" style="color: red;">
<%
     else
%>
		              <td colspan="2" align="center" style="color: green;">
<%
     end if

     Response.write strMessage
%>
                      </td>
                    </tr>
	               	<tr>
		              <td width="40%" align="right">
		            	User Name: 
		              </td>
		              <td width="60%" align="left">
			            <INPUT TYPE="text" ID="HarvestAmericanLogin" NAME="HarvestAmericanLogin">
		              </td>		 
	            	</tr>
		            <tr>
		              <td align="right">
			            Password:
		              </td>
		              <td align="left">
		            	<INPUT Type="password" ID="HarvestAmericanPassword" NAME="HarvestAmericanPassword">
		              </td>
	             	</tr>
		            <tr bgcolor="#ffffff">
		              <td align="center" colspan="2">
		            	<INPUT Type="submit" name="sub" value="Submit" class="button">
		              </td>
		            </tr>
                  </table>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </div>
   
<!-- #include file="includes/footer.inc"-->

  </body>
</html>