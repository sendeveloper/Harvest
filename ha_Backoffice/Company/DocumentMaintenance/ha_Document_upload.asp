<!--#include virtual="includes/connection.asp"-->

<%
    Response.buffer=true
    Response.clear

    Title = "Document Upload"

	If isnull(Request("ID")) or Request("ID") = "" or Request("ID") = "0" Then
		fName = ""
		SubTitle = "Upload a New Document"
	Else
		Dim rs: Set rs = Server.CreateObject("ADODB.Recordset")
		
		strSQL = "SELECT [FileName] FROM ha_Documents WHERE ImageID = " & Request("ID")
		rs.Open strSQL, connCasper10, 1, 3
		
		If Not rs.eof Then
			fName = rs("FileName")
			SubTitle = "Upload a New Version of " & fName
		Else
			fName = ""
			SubTitle = "Upload a New Document"
		End If
		
		rs.Close
	End If
			
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

  <script type="text/javascript">
  
	var fName = '<%=fName%>';
	
    function clickUpload()
        {
		var fPath = document.frm.File1.value;
		
		//Check for no filename
		if (fPath == '')
			{
			alert('No file selected');
			return;
			}
			
		//Check for filenames not matching
		if (fName != '')
			{
			fPath = fPath.replace(/^.*[\\\/]/, '');		
			if (fPath != fName)
				{
				alert('The name of the file you selected does match the document');
				return;
				}
			}
			
		document.frm.submit();
        }
		
	</script>
	
</head>

<%
    Session("refID") = Request("ID")
%>

<body onLoad="SetScreen(600,400);">

  <span class="popupHeading"><%=Title%></span>

<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td>

	  <table style="width: 90%; margin: 0 auto;" cellpadding="2" cellspacing="2">

		<tr><td>&nbsp;</td></tr>
	  
        <tr>
          <td style="text-align: center; font-weight: bold;">
            <%=SubTitle%>
          </td>
        </tr>

		<tr><td>&nbsp;</td></tr>

		<form METHOD="POST" ENCTYPE="multipart/form-data" ACTION="ha_Document_UploadWrite.asp" Name="frm">

        <tr>
          <td>
            <b>Select a file to upload:</b><br>
			<INPUT TYPE=FILE SIZE=50 NAME="File1">
	      </td>
	    </tr>

		</form>

		<tr><td>&nbsp;</td></tr>

      </table>	

    </td>
  </tr>
</table>

    <span class="popupButtons">
      <a href="javascript:clickUpload();" class="bo_Button80">Upload</a>
      <a href="javascript:window.close();" class="bo_Button80">Cancel</a>
    </span>

</body>
</html>
