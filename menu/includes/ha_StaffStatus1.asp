
<%

 requestee = Request("requestFrom")
if requestee <> "" or requestee= "AJAXPage" Then
%>
<!--#include virtual="includes/connection.asp"-->
<%End If%>

<div id="contentFull">
  <div id="columnFull">
    
    <h3 style="padding-top:11px">Staff Status</h3>
<hr/>
<%
			
    set rs = server.createObject("ADODB.Recordset")
    SQL="ha_Menu_StaffStatus" 
    rs.open SQL,connLocal, 3, 3, 4

    If Not rs.eof Then
        While not rs.eof
%>

    <table width="100%" border="0" cellspacing="2" cellpadding="0" class="no-style">
      <tr>
        <td width="12%" align="left">
<%
            If cint(Session("ha_EmployeeID")) = rs("EmpID") Then
%>
          <a href="Javascript:clickStatusChange();"><b><%=rs("FirstName")%></b></a>
<%
            Else
%>
          <b><%=rs("FirstName")%></b>
<%
            End If
%>

        </td>
        <td width="35%" align="left">
          <%=rs("When")%>
        </td>
        <td width="8%" align="center">
          <%=rs("Status")%>
        </td>
        <td width="55%" align="left">
          <%=rs("Note")%>
        </td>
      </tr>
    </table>

<%
            rs.MoveNext
        Wend
    End If

    rs.Close
%>


    <hr>     
  </div>
</div>

