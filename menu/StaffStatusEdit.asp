<!DOCTYPE html>

<!--#include virtual="includes/connection.asp"-->

<head>
    <title>Staff Status</title>

    <meta http-equiv="content-type" content="text/html; charset=utf-8" />

    <link href="<%=strPathMenuIncludes%>ha_menu.css" type="text/css" rel="stylesheet" />
    <!--[if lt IE 7]><script type="text/javascript" src="includes/js/iehoverfix.js"></script><![endif]-->

    <script language="JavaScript1.2" src="includes/js/functions.js"></script>

    <style type="text/css">
	td.inputHeading
		{
		font-weight: bold;
		text-align: right;
		}

    </style>

</head>

<%
    PopupHeading = "Staff Status Edit"

    Dim rs
    set rs = server.createObject("ADODB.Recordset")

    If Request("sMode")="Write" then
        SQL = "ha_Menu_StaffStatus_write(" & Session("ha_EmployeeID") & ", " & _
          Request("ssStatus") & ", " & _
          "'" & Replace(Request("ssNote"),"'","''") & "')"
        Response.write SQL
        'Response.end
        connCasper10.Execute(sql)
	Response.Redirect "includes/PopUpCloser.html"
    Else
      If Request("EmployeeID") = "" Then
        SQL = "ha_Menu_StaffStatus_read(" & Session("ha_EmployeeID") & ")"
      Else
        SQL = "ha_Menu_StaffStatus_read(" & Request("EmployeeID") & ")"
      End If
      'Response.Write(SQL)
      rs.open SQL, connCasper10, 3, 3, 4
    End If
%>

<body onLoad="SetScreen(600, 500);">

  <form Method="Post" Action="StaffStatusEdit.asp?sMode=Write" Name="frm">
  <!--#include virtual="menu/includes/ha_header_popup.inc"-->


  <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr>
      <td>

        <table width="100%" border="0" cellspacing="10" cellpadding="0" align="center">

          <tr>
            <td width="20%">
              &nbsp;
            </td>
            <td width="80%">
              &nbsp;
            </td>
          </tr>

          <tr>
            <td class="inputHeading">
              Name:
            </td>
            <td>
              <%=rs("Name")%>
            </td>
          </tr>


          <tr>
            <td class="inputHeading">
              Last Change:
            </td>
            <td>
              <%=rs("When")%>
            </td>
          </tr>

          <tr><td>&nbsp;</td></tr>

          <tr>
            <td class="inputHeading">
              Status:
            </td>
            <td>
              <select name="ssStatus">
                <%If rs("Status") = "In" or rs("Status") = "Out" Then Selected = "" Else Selected = " Selected" End If%>
                <option value = ""<%=Selected%>>
                <%If rs("Status") = "In" Then Selected = " Selected" Else Selected = "" End If%>
                <option value = "1"<%=Selected%>>In
                <%If rs("Status") = "Out" Then Selected = " Selected" Else Selected = "" End If%>
                <option value = "2"<%=Selected%>>Out
              </select>
            </td>
          </tr>

          <tr>
            <td class="inputHeading">
              Note:
            </td>
            <td>
              <INPUT Type="Text" Name="ssNote" 
                Value="<%=rs("Note")%>"
                Size="60">
            </td>
          </tr>

        </table>



        <table width="100%" border="0" cellspacing="5" cellpadding="5">
          <tr><td>&nbsp;</td></tr>
          <tr>
            <td Width="50%" align="right">
              <a href="javascript:document.frm.submit();" class="button">
                Submit</a> 
            </td> 
            <td Width="50%" align="left">
              <a href="javascript:window.close();" class="button">
                Cancel</a> 
            </td>
          </tr>
          <tr><td>&nbsp;</td></tr>
        </table>

      </td>
    </tr>
  </table>

  <!--#include virtual="menu/includes/ha_footer_popup.inc"-->

  </form>

</body>
</html>
