
   <!--#include virtual="includes/connection.asp"-->
   
<%   
'Declare Variables..
    Dim rs
    Dim PhotoID,strSQL
    
   PhotoID = Request("PhotoId")
   If PhotoID = "" Then PhotoID = 0
   
   'Instantiate Objects
   Set rs = Server.CreateObject("ADODB.Recordset")
   
   
   'Get the specific image based on the ID passed in a querystring
    strSQL = "ha_Document_read(" & PhotoID & ")"
    rs.Open strSQL, connCasper10, 3, 3, 4	
	
    if rs.eof then 'No records found
        Response.End
    else 'Display the contents
		'Response.AddHeader "Content-transfer-encoding", "binary"
		'Response.ContentType = "application/vnd.xls"
		response.Buffer = true
		response.Clear
		'Response.ContentType = "application/octet-stream"
		'Response.ContentType = "application/wordpad"
		'Response.AddHeader "Content-Disposition", "attachment;filename=" & Trim(rs("filename"))
		'Response.AddHeader "Content-Length", rs("filesize")
		'response.CacheControl = "no-store"
        'Response.ContentType = "text/richtext" 
        'Response.ContentType = "application/octet-stream"
		
		'Response.ContentType = "application/vnd.ms-word"
		'Response.AddHeader "Content-transfer-encoding", "binary"
        'Response.AddHeader("content-disposition", "attachment;filename=Tr.doc") 
		
		Response.Addheader "Content-Disposition", "inline; filename=" & Trim(rs("filename"))
		Response.CacheControl = "public"
		'Response.AddHeader "Content-Disposition", "inline"
		'Response.AddHeader "Content-Disposition", "attachment;filename=" & Trim(rs("filename"))
		'Response.AddHeader "Content-Length", rs("filesize")
		Response.ContentType = "application/pdf"
        Response.BinaryWrite(rs("FileData")) 
		
		'Response.ContentType = "application/octet-stream"
		' let the browser know the file name
		'Response.AddHeader "Content-Disposition", "attachment;filename=" & Trim(rs("filename"))
		' let the browser know the file size
		'Response.AddHeader "Content-Length", rs("filesize")
		'Response.BinaryWrite rs("filedata")
    end if
   
	'destroy the variables.
	rs.Close
	connCasper10.Close
	set rs = Nothing
	set connCasper10 = Nothing
   
	'response.Flush
	'response.End
   
 %>