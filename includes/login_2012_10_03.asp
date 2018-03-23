<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<!--#include virtual="includes/connection.asp"-->
<%
    ThisPage = "https://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL")
    
    if Request("HarvestAmericanLogin") <> "" then
        Session("ha_login") = Request("HarvestAmericanLogin")
    end if

    if Request("HarvestAmericanPassword") <> "" then
        Session("ha_password") = Request("HarvestAmericanPassword")
    end if


    if Request("HarvestAmericanLogin") <> "" then
        Session("ha_login") = Request("HarvestAmericanLogin")
    end if

    if Request("HarvestAmericanPassword") <> "" then
        Session("ha_Password") = Request("HarvestAmericanPassword")
    end if

    'response.write "|ha_Login" & Session("ha_Login") & "|<br>"
    'response.write "|ha_Password" & Session("ha_Password") & "|<br>"

%>

<%
  Dim rs
  Set rs=server.createObject("ADODB.Recordset")

  Dim AccountID
  Dim foundCookie
  foundCookie = False
  AccountID = "NULL"
  If Request.Cookies("haBackoffice")("AccountID") <> "" Then
    AccountID = Request.Cookies("haBackoffice")("AccountID")
    foundCookie = True
  End If
  
  'Response.Write("Cookie = [" & Request.Cookies("haBackoffice")("AccountID") & "]")
  'Response.Write("HarvestID = [" & AccountID & "]<br>")
  'Response.Write("Login = [" & fixaps(Request("HarvestAmericanLogin")) & "]<br>")
  'Response.Write("Password = [" & fixaps(Request("HarvestAmericanPassword")) & "]<br>")
  'Response.Write("Referrer = [" & fixaps(Request("ReturnPage")) & "]<br>")

  sql = "ha_BackOffice.dbo.ha_login '" & fixaps(Request("HarvestAmericanLogin")) & "', '" & fixaps(Request("HarvestAmericanPassword")) & "', " & AccountID
  'Response.Write(sql)
  rs.open sql, connDallas, 2, 3
 
  Dim strMessage
  Dim strMessageColorRed
  strMessage = rs("Error")
  'Response.Write("["&strMessage&"]")
  'Response.End

  If isNull(strMessage) Then
    strMessageColorRed = False
    If Not foundCookie Then
      Response.Cookies("haBackoffice")("AccountID") = rs("HarvestID")
      Response.Cookies("haBackoffice").Expires = #December 31, 2030#
    End If
    Session("LoggedIn") = True
    AccountID = rs("HarvestID")
    Session("Status") = rs("Status")
    Session("emp_fname") = rs("FirstName")
    Session("lAccountID") = rs("EmpID")
    Session("login") = rs("login")
    Session("Authority") = rs("Authority")
    Session("ha_fname") = rs("FirstName")
    Session("ha_LastName") = rs("LastName")
    Session("ha_login") = rs("Login")
    'Session("ha_password") = rs("Password")
    Session("ha_EmpID") = rs("EmpID")
	Session("ha_EmpName") = rs("FirstName") & " " & left(rs("LastName"), 1) & "."
    Session("ha_bo_Authorization") = rs("ha_bo_Authorization")
    Session("ha_Permissions") = rs("ha_Permissions")
    rs.Close
    Set rs = Nothing

    connDallas.Execute("exec ha_UserTracking_add '" &_
      Session.SessionId & "', '" &_
      Session("ha_login") & "', '" &_
      Request.ServerVariables("URL") & "',  '" &_
      Request.ServerVariables("REMOTE_ADDR") & "',  '" &_
      Request.ServerVariables("HTTP_REFERER") & "',  '" &_
      Request.ServerVariables("REMOTE_HOST") & "', '" &_
      Request.ServerVariables("LOCAL_ADDR") & "', " &_
      0 & ", " &_
      "NULL" & ", " &_
      2 & ", '" &_
      fixaps(Request("HarvestAmericanLogin")) & "', '" &_
      fixaps(Request("HarvestAmericanPassword")) & "'")

    'Response.Write("<br>")
    'Dim field
    'For Each field in rs.Fields
    '  Response.Write (field.name & ": [" & field & "]<br>")
    'Next
    'Response.End

    If Session("Redirect") = "" Then
      If Session("Redirect") = "" Then
        Response.Redirect strPathBase & "/menu/"
      Else
        Response.Redirect strPathBase & Session("redirect")
      End If    
    Else
      Response.redirect Session("Redirect")
    End If
  Else
   rs.Close
   Set rs = Nothing
   strMessageColorRed = True
   If strMessage = "No credientials nor accountid supplied." Then
     If Request("HarvestAmericanPassword") = "" Then
       strMessage = ""
     Else
       strMessage = "Please enter a username."
     End If
   End If
   If strMessage <> "" Then
     connDallas.Execute("exec ha_UserTracking_add '" &_
       Session.SessionId & "', '" &_
       Session("ha_login") & "', '" &_
       Request.ServerVariables("URL") & "',  '" &_
       Request.ServerVariables("REMOTE_ADDR") & "',  '" &_
       Request.ServerVariables("HTTP_REFERER") & "',  '" &_
       Request.ServerVariables("REMOTE_HOST") & "', '" &_
       Request.ServerVariables("LOCAL_ADDR") & "', " &_
       0 & ", " &_
       "NULL" & ", " &_
       3 & ", '" &_
       fixaps(Request("HarvestAmericanLogin")) & "', '" &_
       fixaps(Request("HarvestAmericanPassword")) & "'")
    End If   
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
  <form method=get action="login.asp" name="frm">
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
<div style="clear: both;">&nbsp;</div>
</div>

<!--#include virtual="menu/includes/ha_footer.inc"-->

</body>
</html>
<%
  End If

Function fixaps (string)
  fixaps = replace(string, "'", "''")
End Function
%>
