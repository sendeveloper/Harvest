<!DOCTYPE html>

<%
    ThisPage = "https://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL")
    
	CRM_Redirect = RequestParse(Request("CRM_Redirect"),"")	
	
    if Request("HarvestAmericanLogin") <> "" then
        Session("ha_Login") = Request("HarvestAmericanLogin")
    end if

    if Request("HarvestAmericanPassword") <> "" then
        Session("ha_Password") = Request("HarvestAmericanPassword")
    end if


    if Request("lname") <> "" then
        Session("ha_Login") = Request("lname")
    end if

    if Request("pass") <> "" then
        Session("ha_Password") = Request("pass")
    end if

    'response.write "|ha_Login=" & Session("ha_Login") & "|<br>"
    'response.write "|ha_Password=" & Session("ha_Password") & "|<br>"
	'response.write "|CookieID=" & Request.Cookies("ha_Backoffice")("ha_EmployeeID") & "|<br>"
	'If Request("HarvestAmericanLogin") <> "" Then
	'	Response.End
	'End If
%>

<!--#include virtual="includes/connection.asp"-->

<%
    '----- Check for login match -----
    messageColorRed = "False"

    if Session("LoggedIn") <> "True" then
	
		set rsLogin=server.createObject("ADODB.Recordset")
		Dim LoginLookup
		LoginLookup = "Valid Login"
		
		'If Session("ha_Login") = "" then 
		'Let's try using the cookie if we're already logged in today
		If Request.Cookies("ha_Backoffice")("ha_EmployeeID") <> "" Then
			eID = Request.Cookies("ha_Backoffice")("ha_EmployeeID")
			r = UserTrackingAdd("Login.asp", "Cookie Found eID: " & eID, "")
		    rsLogin.open "ha_Employee_login_by_id(" & eID & ")",connLocal, 3, 3, 4
			If not rsLogin.eof Then
				If Session("ha_Login") = "" Then 
					Session("ha_Login") = rsLogin("ha_Login")
					LoginLookup = "Auto Login"
				End If
				If Session("ha_Password") = "" Then 
					Session("ha_Password") = rsLogin("ha_Password")
					LoginLookup = "Auto Login"
				End If
			End If
			rsLogin.close
		End If
		'End If
		
		if Session("ha_Login") = "" then 
			Session("ha_Password") = ""
			strMessage = "Please Login"
		elseif Session("ha_Password") = "" then
			Session("ha_Login") = ""
			strMessage = "You forgot to enter a password!"
		else
			SQL="select * from ha_EmpAccounts where " & _
				"username='" & Session("ha_Login") & "'" & _
				"and password='" & Session("ha_Password") & "' " & _
				"AND ActiveStatus='True'"
			rsLogin.open SQL, connCasper10, 2, 3
			
			Session("Redirect") = ""

			if rsLogin.EOF then
				r = UserTrackingAdd("Login.asp", "Invalid Login eID: " & cStr(eID), LoginLookup & Session("ha_Login") & "/" & Session("ha_Password"))
				Session("LoggedIn") = ""
				Session("ha_Login") = ""
				Session("ha_Password") = ""
				strMessage = "Invalid username and/or password!"
				messageColorRed = "True"
			else
				r = UserTrackingAdd("Login.asp", LoginLookup & " eID: " & cStr(eID), Session("ha_Login") & "/" & Session("ha_Password"))
				Session("LoggedIn") = "True"
				Session("ha_fname") = trim(rsLogin("FirstName"))
				Session("ha_lname") = trim(rsLogin("LastName"))
				Session("ha_Login") = trim(rsLogin("username"))
				Session("ha_shortname") = trim(rsLogin("FirstName")) & " " & trim(left(rsLogin("LastName") & " ", 1)) & "."
				Session("ha_Password") = trim(rsLogin("password"))
				Session("ha_EmployeeID") = trim(rsLogin("ID"))
				Session("ha_EmpID") = trim(rsLogin("EmpID"))
				Session("ha_bo_Authorization") = trim(rsLogin("ha_bo_Authorization"))
				Session("ha_Permissions") = trim(rsLogin("ha_Permissions"))
				
				'Drop a cookie
				Response.Cookies("ha_Backoffice")("ha_EmployeeID") = Session("ha_EmployeeID")
				ExpDate = DateAdd("d",5,date)
				Response.Cookies("ha_Backoffice").Expires = ExpDate
				'Response.Cookies("ha_Backoffice").Expires = #12/31/2020#
			end if

			rsLogin.close
        end if
	end if
	
    'response.write "|Redirect=" & Session("Redirect") & "|<br>"
	'response.end
	
    If Session("LoggedIn") = "True" Then
		If CRM_Redirect <> "" Then
			Response.Redirect CRM_Redirect & "?eID=" & Session("ha_EmployeeID")
		End If
        If Session("Redirect") = "" Then
			If Session("ha_EmployeeID") = 1 Then
				Response.Redirect strPathBase & "/menu/dave"
			Else
				Response.Redirect strPathBase & "/menu/"
			End If
        Else
            Response.Redirect strPathBase & Session("Redirect")
        End If
    end if    

%>

<html>
<head>

	<meta http-equiv="content-type" content="text/html; charset=utf-8" />
    <title>Harvest American, Inc. | Login</title>
	<meta name="keywords" content="" />
	<meta name="description" content="" />
    <link href="/menu/includes/ha_menu.css" type="text/css" rel="stylesheet">

</head>

<body onLoad="document.forms.frm.HarvestAmericanLogin.focus()">
  <div  style="padding-top: 70px;">&nbsp;</div>
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
			    <INPUT TYPE="text" NAME="HarvestAmericanLogin" ID="HarvestAmericanLogin" style="width:150px">
		      </td>		 
		    </tr>
	     	<tr>
		      <td align="right">
			  Password:
		      </td>
		      <td align="left">
		      <INPUT Type="password" NAME="HarvestAmericanPassword" ID="HarvestAmericanPassword" style="width:150px">
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

<%
	Function RequestParse(r, d)
		If isnull(r) or r = "" or instr(r,"'") Then
			RequestParse = d
		Else
			RequestParse = r
		End If
	End Function
%>
