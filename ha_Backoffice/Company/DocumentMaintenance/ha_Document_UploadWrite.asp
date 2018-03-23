<!--#include virtual="includes/connection.asp"-->

<%
    Response.buffer=true
    Response.clear

    Title = "File Saved"

	Dim rs: Set rs = Server.CreateObject("ADODB.Recordset")
			
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

<html>
<head>
    <title><%=Title%></title>
    <META http-equiv="Content-Type" content="text/html; charset=UTF-8">
    <META http-equiv="Cache-Control" content="no-cache">


    <link href="/ha_Backoffice/includes/ha_BackOfficePopup.css" type="text/css" rel="stylesheet">
    <script language="javascript" type="text/javascript" src="/includes/haFunctions.js"></script>

</head>

<body onLoad="SetScreen(600,400);">

  <span class="popupHeading"><%=Title%></span>

<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td align="center">

<%

	Response.write Request("File1") & "<br> "
	
    Set Upload = Server.CreateObject("Persits.Upload.1") 

	Upload.OverwriteFiles = True

	'Upload.SetMaxSize 1048576 ' Limit files to 1MB
	'Upload.SetMaxSize 2097152 ' Limit files to 2MB
	'Upload.SetMaxSize 5242880 ' Limit files to 5MB
	Upload.SetMaxSize 10485760 ' Limit files to 10MB
	Count = Upload.Save("c:\upload")

	
    ' Obtain file object
    Set File = Upload.Files("File1") 

    If Not File Is Nothing Then

        fName = File.FileName
        fSize = File.Size
        fType = File.ContentType

		DSN = "driver=SQL Server;server=66.119.55.118,7043;uid=davewj2o;pwd=get2it;database=ha_BackOffice"  'Casper10
		
        SQL = "INSERT INTO ha_Documents" & _
            "(filedata, filename, contenttype, filesize, CreatedBy, CreatedDate) " & _
            "VALUES" & _
            "(?, '" & fName & "', " & _
            "'" & fType & "', " & _
            fSize & ", " & _
            "'" & Session("ha_shortname") & "', " & _
            "'" & now() & "')" 
			
		'Response.write SQL
		
        ' Save to database
        File.ToDatabase DSN, SQL
		
		'Index file
		SQL = "ha_Document_index('" & fName & "')"
		connCasper10.Execute(SQL)
		
    End If

	
%>


      <table width="100%" border="0" cellspacing="5" cellpadding="5">

		<tr><td>&nbsp;</td></tr>
		
        <tr>
          <td width="40%" align="right">
            File Name:
          </td>
          <td width="60%" align="left">
            <b><%=fName%></b>
          </td>
        </tr>

        <tr>
          <td align="right">
            File Size:
          </td>
          <td align="left">
            <b><%=FormatNumber(fSize, 0)%></b>
          </td>
        </tr>

        <tr>
          <td align="right">
            File Type:
          </td>
          <td align="left">
            <b><%=fType%></b>
          </td>
        </tr>

		<tr><td>&nbsp;</td></tr>
      </table>	

	  
    <span class="popupButtons">
      <a href="javascript:window.close();" class="bo_Button80">Close</a>
    </span>

    </td>
  </tr>
</table>
	
<script>
    window.opener.location.href = window.opener.location.href;
</script>

</body>
</html>

