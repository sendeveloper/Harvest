 <!--#include virtual="includes/connection.asp"-->

<html>
<head>
    <title>E-mail Editor</title>

    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <script language="JavaScript" src="/ha_Backoffice/includes/hafunctions.js"></script>

    <link href="../includes/BackOfficePopup.css" type="text/css" rel="stylesheet">

<style type="text/css">


td
	{
	border-width:	0;
	font-size:		11pt;
	}

</style>

</head>

<%
    If Request("ID") = "" or isnull(Request("ID")) then
        ID = 0
    Else
        ID = Request("ID")
    End If

    Dim rs
    Dim SQL

    set rs = server.createObject("ADODB.Recordset")

    SQL = "SELECT * FROM ha_Emails.dbo.ha_EmailEditor WHERE EditorName = '" & Request("e") & "'"
    rs.open SQL, connCasper10, 1, 2

    Title = rs("EditorHeading")
    StoredProcedure = rs("StoredProcedure")

    rs.close

    SQL = "ha_Emails.dbo." & StoredProcedure & "(" & ID & ", '" & Session("login") & "')" 
    rs.open SQL, connCasper10, 3, 3, 4

    eSendTo = rs("Email")
    eSendFrom = rs("SendFrom")
    eSubject = rs("Subject")
    eText = rs("eText")

    rs.close
    set rs=nothing
%>


<FORM METHOD="Post" ACTION="ha_Email_Send.asp" name="frm">
<body onLoad="SetScreen(1100,800);">
	<span class="popupHeading"><%=Title%></span>

<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center" style="padding-top: 10px;">
  <tr>
    <td colspan="2">
      <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
          <td>

			<table width="100%" border="0" cellspacing="3" cellpadding="0" align="center" style="border: 1px solid black; background-color: #B2E5FF;">

              <tr>
                <td align="center" colspan="2">
                  This E-mail will be sent from 
                  <INPUT TYPE="Text" NAME="SendFrom" ID="SendFrom" Value="<%=eSendFrom%>" Size="30">
                </td>		 
              </tr>

              <tr>
                <td width="10%" align="right">
                  Send To
                </td>
                <td width="80%">
                  <INPUT TYPE="Text" NAME="SendTo" ID="SendTo" Value="<%=eSendTo%>" Size="110">
                </td>		 
              </tr>

              <tr>
                <td align="right">
                  Copy
                </td>
                <td>
                  <INPUT TYPE="Text" NAME="Copy" ID="Copy" Value="" Size="110">
                </td>		 
              </tr>

              <tr>
                <td align="right">
                  Subject
                </td>
                <td>
                  <INPUT TYPE="Text" NAME="Subject" ID="Subject" Value="<%=eSubject%>" Size="110">
                </td>		 
              </tr>

              <tr>
                <td align="right" valign="top">
                  Text
                </td>
                <td>
                  <textarea rows="25" cols="90" Name="noteTextArea" ID = "noteTextArea"><%=eText%></textarea>
                </td>		 
              </tr>

			  <tr><td>&nbsp;</td></tr>
			  
            </table>

		
          </td>
        </tr>
      </table>

      <table width="100%" border="0" cellspacing="2" cellpadding="2" style="padding-top: 10px;">
        <tr>
          <td width="50%">&nbsp;</td>
          <td>
            <a href="javascript:document.frm.submit();" class="buttons bo_Button80">Send</a>
          </td>
          <td>
            <a href="javascript:window.close();" class="buttons bo_Button80">Cancel</a>
          </td>
          <td width="50%">&nbsp;</td>
        </td>
      </table>

    </td>
  </tr>
</table>


</body>
</form>
</html>
