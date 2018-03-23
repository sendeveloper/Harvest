<!--#include virtual="includes/connection.asp"-->
   
<%   
	Response.Buffer = true
	Response.Clear

	'Declare Variables..
    Dim rs
    Dim PhotoID, strSQL, strContentType
    
	PhotoID = Request("PhotoId")
	If PhotoID = "" Then PhotoID = 0
   
	'Instantiate Objects
	Set rs = Server.CreateObject("ADODB.Recordset")
   
   
	'Get the specific image based on the ID passed in a querystring
    strSQL = "ha_Document_read(" & PhotoID & ")"
	'response.write strSQL & "<br>"
    rs.Open strSQL, connCasper10, 3, 3, 4	
	strContentType = rs("ContentType")
	strFileName = trim(rs("FileName"))
	strFileSize = trim(rs("FileSize"))
	rs.close
	
	
	If strContentType = "" Then
		'No record found
        Response.End
    Else 		
		Select Case strContentType
		Case "image/jpeg"
			strContentType = "image/jpeg"
		Case "application/pdf"
			strContentType = strContentType
		Case Else
			'strContentType = ""
		End Select
		
		If strContentType > "" Then
			'Show it if you go it
			strSQL = "ha_Document_read(" & PhotoID & ", 1)"
			rs.Open strSQL, connCasper10, 3, 3, 4	

			'Response.Addheader "Content-Disposition", "inline; filename=" & strFileName
			Response.Addheader "Content-Disposition", "inline"
			Response.AddHeader "Content-Length", strFileSize
			Response.CacheControl = "public"
				
			Response.ContentType = strContentType
			Response.BinaryWrite(rs("FileData"))
		End If
    End If
   
	'Destroy the variables.
	rs.Close
	connCasper10.Close
	set rs = Nothing
	set connCasper10 = Nothing
   
   	Response.Flush
	Response.End

   
 %>