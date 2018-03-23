<%
    if Session("Redirect") = "" AND Request.ServerVariables("PATH_INFO") <> "/includes/login.asp" then
        Session("Redirect") = Request.ServerVariables("PATH_INFO")
    end if
'response.write "|" & Session("Redirect") & "|<br>"
'response.write "|" & Request.ServerVariables("HTTPS") & "|<br>"

'response.write "|" & Request("lname") & "|<br>"
'response.write "|" & Request("pass") & "|<br>"

    Session.Timeout = 120

    '----- Clear Browser Cache (IE & FF) -----
    Response.AddHeader "pragma","no-cache"
    Response.AddHeader "cache-control","private"
    Response.AddHeader "Pragma", "no-store"
    Response.CacheControl = "no-store"

    '----- Directories -----
    strPathBase         =      ""
    strPathIncludes     =      strPathBase & "/includes/"
    strPathControlPanel =      strPathBase & "/controlpanel/"
    strPathMenuIncludes =      strPathBase & "/menu/includes/"
    strPathMenuImages   =      strPathBase & "/menu/images/"
    strPathMobile       =      strPathBase & "/mobile/"

    '----- Be sure secure path is used -----
'    if Request.ServerVariables("HTTPS") <> "on" then
'        if UCase(Session("Redirect")) = "/MOBILE/MLOGIN.ASP" then
'            response.redirect "" & Session("Redirect")
'        else
'            response.redirect "" & Session("Redirect")
'        end if
'    end if

    '----- Check for login status -----
    if Request("lname") = "nopassword" then
        'Specifically asking for the login screen
        Session("LoggedIn") = ""
    end if


'    if Session("LoggedIn") = "True" then 
'        response.write "|" & Session("lname") & "|<br>"
'        response.write "|" & Session("ha_login") & "|<br>"
        'response.end
'    end if


    '----- Open Database -----
    dim objcon

    set objcon=server.CreateObject("ADODB.Connection")

    objcon.Open "driver=SQL Server;server=dallas01.harvestamerican.net;uid=davewj2o;pwd=get2it;database=ha_BackOffice"

    '----- User Tracking -----

    Dim URL
    Dim Domain
    Dim p

    URL = Request.ServerVariables("HTTP_REFERER")
    if len(URL) > 250 then
        URL = Left(URL, 250)
    end if

    p = instr(URL,"//")
    If p > 0 then
        Domain = Right(URL,Len(URL) - (p + 1))
    Else
        Domain = URL
    End If

    p = instr(Domain,"/")
    If p > 0 then
        Domain = Left(Domain,p - 1)
    End If

    sql = "ha_add_UserTracking('" & Session.SessionID & "', " & _
        "'" & Session("ha_login") & "', " & _
        "'" & Request.ServerVariables("PATH_INFO") & "', " & _
        "'" & Request.ServerVariables("REMOTE_ADDR") & "', " & _
        "'" & URL & "', " & _
        "'" & Domain & "', " & _
	    "'', " & _
        "'')"

    objcon.Execute(sql)

'response.write "|" & Session("ha_login") & "|<br>"

    if Session("ha_login") = "" then
        'We're not logged in
        if UCASE(Request.ServerVariables("PATH_INFO")) <> "/INCLUDES/LOGIN.ASP" and UCASE(Request.ServerVariables("PATH_INFO")) <> "/MOBILE/MLOGIN.ASP" then
            'Must be we weren't logging in either.  Let's go to login page
            if UCASE(Request.ServerVariables("PATH_INFO")) = "/MOBILE/MLOGIN.ASP" or _
               instr(Request.ServerVariables("HTTP_USER_AGENT"),"iPhone") then
               Response.Redirect strPathMobile & "mlogin.asp"
            else
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

                Response.Redirect strPathIncludes & "login.asp"
            end if
        end if
    end if

%>
 
