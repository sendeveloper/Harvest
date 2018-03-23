<%
    Dim rs
    Dim SQL
    Dim connLocal

    set connLocal=server.CreateObject("ADODB.Connection")

    connLocal.Open "driver=SQL Server;server=127.0.0.1,7043;uid=davewj2o;pwd=get2it;database=ha_Backoffice"	 
    set rs = server.createObject("ADODB.Recordset")

    SQL = "SELECT * from ha_Help " & _
        "WHERE [ID] = " & Request("ID")
	'Response.write SQL

    rs.open SQL, connLocal, 2, 3

    response.ContentType = "text/xml"
    response.write("<?xml version='1.0' encoding='ISO-8859-1'?>")

    if rs.eof then
        response.write("<Help>")
        response.write("<HelpResponseTitle>" & "error" & "</HelpResponseTitle>" )
        response.write("</Help>")
    else
        response.write("<Help>")
        response.write("<Top>" & rs("PosTop") & "</Top>" )
        response.write("<Left>" & rs("PosLeft") & "</Left>" )
        response.write("<Width>" & rs("Width") & "</Width>" )
        response.write("<Height>" & rs("Height") & "</Height>" )
        response.write("<HelpResponseReference>" & rs("Reference") & "</HelpResponseReference>" )
        response.write("<HelpResponseTitle>" & rs("Title") & "</HelpResponseTitle>" )
        response.write("<HelpResponseBody>" & fixChars(rs("Body")) & "</HelpResponseBody>" )
        response.write("</Help>")
    end if

    rs.Close
    connLocal.close

	FUNCTION fixChars(ByVal s)
	    s = Replace(s, "&", "&amp;" )
	    s = Replace(s, "'", "&apos;" )
	    s = Replace(s, chr(34), "&quot;" )
	    s = Replace(s, "<", "&lt;" )
	    fixChars = Replace(s, ">", "&gt;" )
	END FUNCTION

%>


