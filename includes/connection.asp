
<%
  If Session("Redirect") = "" And Request.ServerVariables("PATH_INFO") <> "/includes/login.asp" Then
    Session("Redirect") = Request.ServerVariables("PATH_INFO")
  End If

  Session.Timeout = 120

  '----- Clear Browser Cache (IE & FF) -----
  Response.AddHeader "pragma","no-cache"
  Response.AddHeader "cache-control","private"
  Response.AddHeader "Pragma", "no-store"
  Response.CacheControl = "no-store"

  '----- Directories -----
  
  Dim strPathBase:         strPathBase         = "http://www.HarvestAmerican.info"
  Dim strPathIncludes:     strPathIncludes     = strPathBase & "/includes/"
  Dim strPathControlPanel: strPathControlPanel = strPathBase & "/controlpanel/"
  Dim strPathMenu:         strPathMenu         = strPathBase & "/menu/"
  Dim strPathMenuIncludes: strPathMenuIncludes = strPathBase & "/menu/includes/"
  Dim strPathMenuImages:   strPathMenuImages   = strPathBase & "/menu/images/"
  Dim strPathMobile:       strPathMobile       = strPathBase & "/mobile/"

  '----- Be sure secure path is used -----
  If Request.ServerVariables("HTTPS") <> "on" Then
      'strSecureURL = "https://"
      'strSecureURL = strSecureURL & Request.ServerVariables("SERVER_NAME")
      'strSecureURL = strSecureURL & Session("Redirect")
      'Response.Redirect strSecureURL
  End If

  '----- Check for login status -----
  If Request("lname") = "nopassword" Then
    'Specifically asking for the login screen
    Session("LoggedIn") = ""
  End If


    '----- Open Database -----

  Dim connPubTables: Set connPubTables = Server.CreateObject("ADODB.Connection")
  connPubTables.Open "driver=SQL Server;server=66.119.50.230,7043;uid=davewj2o;pwd=get2it;database=ha_PublishedTables"

  Dim connCasper10: Set connCasper10 = Server.CreateObject("ADODB.Connection")
  connCasper10.Open "driver=SQL Server;server=66.119.50.230,7043;uid=davewj2o;pwd=get2it;database=ha_BackOffice"

  Dim connLocal: Set connLocal = Server.CreateObject("ADODB.Connection")
  connLocal.Open "driver=SQL Server;server=127.0.0.1,7043;uid=davewj2o;pwd=get2it;database=ha_BackOffice"
  
  dim objcon:set objcon=server.CreateObject("ADODB.Connection")
  objcon.Open "driver=SQL Server;server=127.0.0.1,7043;uid=davewj2o;pwd=get2it;database=z2t_Backoffice"	 
%>

<!--#include virtual="ha_Backoffice/includes/FunctionsUserTracking.inc"-->
 
 
<%  
  '----- User Tracking -----
  r = UserTrackingAdd("Login.asp", "Connection.asp", "")
  Dim URL
  Dim Domain
  Dim p

  URL = Request.ServerVariables("HTTP_REFERER")
  If len(URL) > 250 Then
    URL = Left(URL, 250)
  End If

  p = instr(URL, "//")
  If p > 0 Then
    Domain = Right(URL,Len(URL) - (p + 1))
  Else
    Domain = URL
  End If

  p = instr(Domain, "/")
  If p > 0 Then
    Domain = Left(Domain,p - 1)
  End If

  sqlText = "ha_UserTracking_add('" & Session.SessionID & "', " & _
    "'" & Session("ha_login") & "', " & _
    "'" & Request.ServerVariables("PATH_INFO") & "', " & _
    "'" & Request.ServerVariables("REMOTE_ADDR") & "', " & _
    "'" & URL & "', " & _
    "'" & Domain & "', " & _
    "'', " & _
    "'')"
  'objcon.Execute(sqlText)
  connLocal.Execute(sqlText)
'response.write "|" & Session("ha_login") & "|<br>"

  If Session("ha_login") = "" Then
    'We're not logged in
    If Lcase(Request.ServerVariables("PATH_INFO")) <> "/includes/login.asp" And _
       Lcase(Request.ServerVariables("PATH_INFO")) <> "/mobile/mlogin.asp" Then
                        
      'Must be we weren't logging in either.  Let's go to login page
      If Lcase(Request.ServerVariables("PATH_INFO")) = "/mobile/mlogin.asp" Or _
        Instr(Request.ServerVariables("HTTP_USER_AGENT"),"iPhone") Or _
		Instr(Request.ServerVariables("HTTP_USER_AGENT"),"Android") Then
          Response.Redirect strPathMobile & "mlogin.asp"
      Else
        If Request("HarvestAmericanLogin") <> "" Then
          Session("ha_login") = Request("HarvestAmericanLogin")
        End If

        If Request("HarvestAmericanPassword") <> "" Then
          Session("ha_password") = Request("HarvestAmericanPassword")
        End If

        If Request("lname") <> "" Then
          Session("ha_login") = Request("lname")
        End If

        If Request("pass") <> "" Then
          Session("ha_Password") = Request("pass")
        End If

        Response.Redirect "/includes/login.asp"
      End If
    End If
  End If

%>
 
