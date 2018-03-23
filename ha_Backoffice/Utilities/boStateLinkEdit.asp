<%
	Response.buffer=true
	Response.clear
	Dim strMessage
	Dim strMode
%>

  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/Config.asp"-->
  <!--#include virtual="ha_Backoffice/Utilities/connection_Barley.asp"-->

<html>
<head>
    <title>Number-it State Regulations Link Edit</title>

    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

    <link href="/ha_BackOffice/includes/datepick/jquery-ui.css" rel="stylesheet" />
	<link href="/ha_BackOffice/includes/ha_Backoffice.css" type="text/css" rel="stylesheet" />
	<link href="/ha_Backoffice/includes/ha_BackOfficePopup.css" type="text/css" rel="stylesheet">

</head>

    <%
    IF session("loggedin")<>"True" OR isNULL(session("loggedin")) THEN
	Response.Redirect "../bologin.asp"
    END IF

    If Request("State")="" or isnull(Request("State")) then
	If Session("State")="" or isnull(Session("State")) then
	    Close
	End If
    Else
	Session("State")=Request("State")
    End If

	Title = "State Regulations Link Edit"
	Dim rs
	Dim SQL
	set rs = server.createObject("ADODB.Recordset")
	SQL="SELECT * FROM ni_StateRegulations " & _
	    "WHERE [Abbreviation] = '" & Session("State") & "'"
		
	rs.open SQL,objcon,2,3
	eLink = rs("URL")
    %>

<FORM METHOD="Post" ACTION="boStateLinkPost.asp" name="frm">
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" 
	marginwidth="0" marginheight="0" 
	link="#333366" vlink="#333366" alink="#3399FF">

<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td colspan="2">
      <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
          <td>
            <table width="100%" border="0" cellspacing="10" cellpadding="0" align="center">

              <tr>
                <td>
                  <BR><Font Size="4"><Center><b>State Regulations Link Edit</Center></b></Font>
                </td>
              </tr>

            </table>	

	    <table width="100%" border="0" cellspacing="10" cellpadding="0" align="center">

		<tr>
		    <td align="right" Width="15%">
		 	<B>Abbreviation:</B>
		    </td>
		    <td Width="85%">
	                <%=rs("Abbreviation")%>
		    </td>		 
		</tr>

		<tr>
		    <td align="right">
		 	<B>Name:</B>
		    </td>
		    <td>
	                <%=rs("Description")%>
		    </td>		 
		</tr>

		<tr>
		    <td align="right">
		 	<B>Country:</B>
		    </td>
		    <td>
	                <%=rs("Country")%>
		    </td>		 
		</tr>

		<tr>
		    <td align="right">
			<B>Link:</B>
		    </td>
		    <td>
			<INPUT TYPE="Text" NAME="Link"
			    Value="<%=eLink%>"
			    Size="80">
	                
		    </td>		 
		</tr>

		<tr>
		     <td Align="center" colspan="2">
			<a href="javascript:document.frm.submit()"class="bo_Button80">Submit</a>
			<!-- <img src="<%=strPathImages%>but_red_submit.gif" border="0"></a>  -->
			<a href="javascript:window.close();" class="bo_Button80">Close</a>
		     </td>
		</tr>

            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>


</body>
</form>
</html>
