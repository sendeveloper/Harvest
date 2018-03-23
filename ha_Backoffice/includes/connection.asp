 
<%
	'----- Open Database -----
	dim objcon, connCasper08, connEmail
	set objcon=server.CreateObject("ADODB.Connection")
	set connCasper08=server.CreateObject("ADODB.Connection")
	set connEmail=server.CreateObject("ADODB.Connection")

 	'objcon.Open "driver=SQL Server;server=68.178.202.54;uid=davewj2o;pwd=get2it;database=ha_prod"
 	objcon.Open "driver=SQL Server;server=66.119.50.228,7843;uid=davewj2o;pwd=get2it;database=ha_prod"
	connEmail.Open 	"driver=SQL Server;server=66.119.50.230,7043;uid=davewj2o;pwd=get2it;database=ha_Emails"
 	connCasper08.Open "driver=SQL Server;server=66.119.50.228,7843;uid=davewj2o;pwd=get2it;database=ha_CRM"

	'----- Function To Display Navigation Bar -----
	sub displayNavigationbar(str)
		percentfound=instr(str,"%")
		if cint(percentfound)=0 then
			colonfound=instr(str,":")
			if cint(colonfound)<>0 then
				secondarray=split(str,":")
				texttoshow=secondarray(0)
				response.write("<table width=""80%""><tr><td>	<font size=""2""  >"&texttoshow&"</font></td></tr></table>")
			end if
		else
			firstarray=split(str,"%")
			response.write("<table width=""80%""><tr><td><font size=""2"">")
			for i=0 to ubound(firstarray)
				secondarray=split(firstarray(i),":")
				texttoshow=secondarray(0)
				hreflink=secondarray(1)
				if trim(cstr(hreflink))="last" then
					response.write(texttoshow)
				else
					response.write("<a href="""&hreflink&"""><font size=""2"">"&texttoshow&"</font></a>&nbsp;&gt;&gt;&nbsp;")
				end if
			next
			response.write("</font></td></tr></table>")
		end if
	end sub

	'----- Function to fix the Quotes -----
	FUNCTION fixQuotes( ByVal theString )
		fixQuotes = Replace( theString, "'", "''" )
	END FUNCTION

      Session("IPBigInt") = IPConvert(Request.ServerVariables("REMOTE_ADDR"))


	'----- Function to look-up Country's IP -----
      FUNCTION findCountry()
          'DIM rs
          'SET rs = Server.CreateObject("ADODB.Recordset")
          'rs.open "ni_read_CountryByBigInt(" & Session("IPBigInt") & ")", objCon, 3, 3, 4

          'if not rs.eof then
          '    Session("CountryAbr3") = rs("Abr3")
          '    findCountry = rs("Name")
          'else
          '    Session("CountryAbr3") = "USA"
          '    findCountry = "UNITED STATES"
          'end if

          'if Session("CountryAbr3") = "ZZZ" then
          '    Session("CountryAbr3") = "USA"
          '    findCountry = "UNITED STATES"
          'end if

          'rs.Close
      END FUNCTION

    SUB readCurrency

		'dim con, str
        'dim rs, SQL

		'set con=server.CreateObject("ADODB.Connection")
        'str="driver=SQL Server;server=68.178.202.54;uid=cbs_web_user;pwd=web_user_cbs;database=ha_prod"

		'con.Open str

        'set rs = Server.CreateObject("ADODB.Recordset")
        'rs.open "ni_read_Currency(" & Session("CountryAbr3") & ")", con, 3, 3, 4

        'if not rs.eof then
        '    Session("Currency") = rs("Name")
        '    Session("Bid") = rs("Bid")
        '    Session("Symbol") = rs("Symbol")
        'else
            Session("Currency") = "United States Dollar"
            Session("Bid") = 1
            Session("Symbol") = "USD"
        'end if
        'rs.Close
        'con.close
        'set con = nothing
    END SUB

	'----- Function to convert IP addresses -----
      Public Function IPConvert(IPAddress)

          Dim x
          Dim Pos
          Dim PrevPos
          Dim Num

          If IsNumeric(IPAddress) Then
              IPConvert = "0.0.0.0"
              For x = 1 To 4
                  Num = Int(IPAddress / 256 ^ (4 - x))
                  IPAddress = IPAddress - (Num * 256 ^ (4 - x))
                  If Num > 255 Then
                      IPConvert = "0.0.0.0"
                      Exit Function
                  End If
      
                  If x = 1 Then
                      IPConvert = Num
                  Else
                      IPConvert = IPConvert & "." & Num
                  End If
              Next
          ElseIf UBound(Split(IPAddress, ".")) = 3 Then
              For x = 1 To 4
                  Pos = InStr(PrevPos + 1, IPAddress, ".", 1)
                  If x = 4 Then Pos = Len(IPAddress) + 1
                  Num = Int(Mid(IPAddress, PrevPos + 1, Pos - PrevPos - 1))
                  If Num > 255 Then
                      IPConvert = "0"
                      Exit Function
                  End If
                  PrevPos = Pos
                  IPConvert = ((Num Mod 256) * (256 ^ (4 - x))) + IPConvert
              Next
          End If

      End Function

	'----- User Tracking -----	
	'Post the page load into activity
    sql = "crm_Activity_add('" &_
      Session("ha_ShortName") & "', " & _
	  "'formLoad', " &_
      "'" & Request.ServerVariables("path_info") & "', " &_
      "'" & Request.ServerVariables("HTTP_REFERER") & "', " &_
	  "'', " & _
	  "'', " & _
	  "0, " & _
      "'" & Session.SessionID & "', " &_
      "'" & Request.ServerVariables("REMOTE_ADDR") & "', " &_
      "'(ni) Connection.asp')"	  

	connCasper08.Execute(sql)
%>
 
