<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<%
    ThisPage = "https://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL")
    
    if Request("HarvestAmericanLogin") <> "" then
        Session("ha_login") = Request("HarvestAmericanLogin")
    end if

    if Request("HarvestAmericanPassword") <> "" then
        Session("ha_password") = Request("HarvestAmericanPassword")
    end if


    if Request("lname") <> "" then
        Session("ha_login") = Request("lname")
    end if

    if Request("pass") <> "" then
        Session("ha_Password") = Request("pass")
    end if

    'response.write "|ha_Login" & Session("ha_Login") & "|<br>"
    'response.write "|ha_Password" & Session("ha_Password") & "|<br>"

%>

<!--#include virtual="includes/connection.asp"-->

<%
    '----- Check for login match -----
    messageColorRed = "False"
    
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
        response.write SQL & "|<br>"
        RS.open SQL,connCasper10,2,3

        if RS.EOF then
	      Session("LoggedIn") = ""
          Session("ha_login") = ""
          Session("ha_password") = ""
          strMessage = "Invalid username and/or password!"
          messageColorRed = "True"
        else
          Session("LoggedIn") = "True"
          Session("ha_fname") = trim(rs("FirstName"))
          Session("ha_lname") = trim(rs("LastName"))
          Session("ha_login") = trim(rs("username"))
          Session("ha_password") = trim(rs("password"))
          Session("ha_EmpID") = trim(rs("EmpID"))
          Session("ha_Permissions") = trim(rs("ha_Permissions"))
          Session("ha_bo_Authorization") = trim(rs("ha_bo_Authorization"))
          if Session("Redirect") = "" then
            .Response.Redirect strPathBase & "/menu/"
          else
            Response.Redirect strPathBase & Session("redirect")
          end if
        end if

        rs.close
        set rs = nothing
    
        if Session("LoggedIn") = "True" then
           if Session("Redirect") = "" then
              Response.Redirect strPathBase & "/menu/"
           else
              Response.Redirect strPathBase & Session("redirect")
           end if
        end if    
    end if
%>

<html xmlns="http://www.w3.org/1999/xhtml">

<head>

	<meta http-equiv="content-type" content="text/html; charset=utf-8" />
    <title>Harvest American, Inc. | Login</title>
	<meta name="keywords" content="" />
	<meta name="description" content="" />
    <link href="<%=strPathMenuIncludes%>ha_menu.css" type="text/css" rel="stylesheet">

</head>

<body onLoad="document.forms.frm.HarvestAmericanLogin.focus()">

<!--#include virtual="menu/includes/ha_header.inc"-->

<!--#include virtual="menu/includes/ha_menu.inc"-->

<div id="content">
<div id="columnA">
  <FORM method=get action="login.asp" name="frm">
    <table width="400" border="1" cellspacing="0" cellpadding="0" align="center" bgColor="#FFFFFF" style="margin-top: 15px;">
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
			    <INPUT TYPE="text" NAME="HarvestAmericanLogin" style="width:150px">
		      </td>		 
		    </tr>
	     	<tr>
		      <td align="right">
			  Password:
		      </td>
		      <td align="left">
		      <INPUT Type="password" NAME="HarvestAmericanPassword" style="width:150px">
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
  </form>
</div>
<div id="columnB">

<!--#include virtual="menu/includes/ha_domains.inc"-->

</div>
<div style="clear: both;">&nbsp;</div>
</div>

<!--#include virtual="menu/includes/ha_footer.inc"-->

</body>
</html>
