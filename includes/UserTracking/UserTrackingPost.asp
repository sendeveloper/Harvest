<%
    Response.buffer=true
    Dim rs
    Dim Data1
    Dim Data2
    Dim URL
    Dim Domain
    Dim p
    Dim connCasper10
    Dim pgeURL

    Data1 = Request("Data1")
    Data2 = Request("Data2")

    set connCasper10=server.CreateObject("ADODB.Connection") 
	connCasper10.Open "driver=SQL Server;server=10.119.55.118,7043;uid=davewj2o;pwd=get2it;database=ha_BackOffice"

      if Request("Page")="" then
          pgeURL = Request.ServerVariables("path_info")
      else
          pgeURL = Request("Page")
      end if

      if left(pgeURL,1) = "/" then
          pgeURL = mid(pgeURL&"  ",2)
      end if

      pgeURL = trim(pgeURL)

	URL = Request.ServerVariables("HTTP_REFERER")

	p = instr(URL,"?s=google")
	If p > 0 then
            Domain = "Google Pay-Per-Click"
        Else
	    p = instr(URL,"//")
	    If p > 0 then
	        Domain = Right(URL,Len(URL) - (p + 1))
	    Else
	        Domain = URL
            End If
	End If

	p = instr(Domain,"/")
	If p > 0 then
	    Domain = Left(Domain,p - 1)
	End If

	Response.write Request.ServerVariables("path_info") & "<br>"
	Response.write pgeURL & "<br>"
	Response.write Request.ServerVariables("REMOTE_ADDR") & "<br>"
	Response.write URL & "<br>"

    SQL = "ha_UserTracking_add('" & Session.SessionID & "', " & _
            "'" & Session("ha_ShortName") & "', " & _
            "'" & pgeURL & "', " & _
            "'" & Request.ServerVariables("REMOTE_ADDR") & "', " & _
            "'" & URL & "', " & _
            "'" & Domain & "', " & _
            "'" & Session("RequestSource") & "', " & _
            "'" & Session("IPBigInt") & "', " & _
            "'', " & _
            "'" & Request("Op") & "', " & _
            "'" & Data1 & "', " & _
            "'" & Data2 & "')"

	Response.write SQL
    connCasper10.Execute (SQL)
    connCasper10.close
%>


