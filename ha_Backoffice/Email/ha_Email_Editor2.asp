<!--#include virtual="ha_backoffice/includes/config.asp"-->
<!--#include virtual="ha_backoffice/includes/connection.asp"-->
<a href="#" title="Close" class="closeReport">X</a>

<%
    IF session("loggedin")<>"True" OR isNULL(session("loggedin")) THEN
	Response.Redirect "../bologin.asp"
    END IF

    If Request("ID") = "" or isnull(Request("ID")) then
        Session("ID") = 0
    Else
        Session("ID") = Request("ID")
    End If

    Dim rs
    Dim SQL

    set rs =  server.createObject("ADODB.Recordset")

    SQL = "SELECT * FROM ha_EmailEditor WHERE EditorName = '" & Request("e") & "'"
''	response.Write(sql)
	
    rs.open SQL, connEmail, 1, 2

    EditorHeading = rs("EditorHeading")
    StoredProcedure = rs("StoredProcedure")

    rs.close

    SQL = StoredProcedure & "(" & Session("ID") & ", '" & Session("login") & "')" 
  ''  Response.Write SQL
    rs.open SQL, connEmail, 3, 3, 4

    eSendTo = rs("Email")
    eSendFrom = rs("SendFrom")
    eSubject = rs("Subject")
    eText = rs("eText")

    rs.close
    set rs=nothing
%>

<a href="#" id="print-report" onclick="javascript:printDiv(&#34modal-Email&#34)" style="text-decoration:none;float:right;"><img src="../images/download.jpg" class="print" style="width: 25px;  height: 25px;"/></a>

<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td colspan="2">
      <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
          <td>
            <table width="100%" border="0" cellspacing="10" cellpadding="0" align="center">

              <tr>
                <td>
                  <Font Size="4"><Center><b><%=EditorHeading%></b></Center></Font> 
                </td>
              </tr>

            </table>	

			<table width="100%" border="0" cellspacing="3" cellpadding="0" align="center" style="border-radius:5px;border: 1px solid black; background-color: #Eaeaea">

              <tbody>
              <tr>
                <td width="25%" align="right">
                <input type = "hidden" value="<%=sql%>" id="hdnSQL" />
                This E-mail will be sent from 
                </td>
                <TD>
                  
                  <input type="Text" name="SendFrom" id="SendFrom" value="<%=eSendFrom%>" size="88">
                </td>		 
              </tr>

              <tr>
                <td width="25%" align="right">
                  Send To
                </td>
                <td width="70%">
                  <input type="Text" name="SendTo" id="SendTo" value="<%=eSendTo%>" size="88">
                </td>		 
              </tr>

              <tr>
                <td align="right">
                  Copy
                </td>
                <td>
                  <input type="Text" name="Copy" id="Copy" value="" size="88">
                </td>		 
              </tr>

              <tr>
                <td align="right">
                  Subject
                </td>
                <td>
                  <input type="Text" name="Subject" id="Subject" value="<%=eSubject%>" size="88">
                </td>		 
              </tr>

              <tr>
                <td align="right" valign="top">
                  Text
                </td>
                <td>
                  <textarea rows="25" cols="90" name="noteTextArea" id="noteTextArea" disabled="disabled"><%=eText%></textarea>
                </td>		 
              </tr>

			  <tr><td>&nbsp;</td></tr>
			  
            </tbody></table>

		
          </td>
        </tr>
      </tbody></table>

      <table width="100%" border="0" cellspacing="2" cellpadding="2" style="padding-top: 10px;">
        <tbody><tr>
          <td width="50%">&nbsp;</td>
          <td>
            <a href="#" onclick="javascript:sendEmail();" class="buttons bo_Button80">Send</a>
          </td>
          <td>
            <a href='javascript:$("#modal-Email").css("display","none");$("#fade-report").css("display", "none");' class="buttons bo_Button80">Cancel</a>
          </td>
          <td width="50%">&nbsp;</td>
        
      </tr></tbody></table>

    </td>
  </tr>
</table>
<div id="email-response"></div>


