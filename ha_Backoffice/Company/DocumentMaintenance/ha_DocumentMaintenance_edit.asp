<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/Config.asp"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsFilter.inc"-->
  
<%
    Title = "Document Maintenance Edit"

	Dim rs: Set rs = Server.CreateObject("ADODB.Recordset")

	strSQL = "ha_Document_info(" & Request("id") & ")"
		
    rs.open strSQL, connCasper10, 3, 3, 4

		eType = rs("DocType")
		eBusiness = rs("Business")
		eDepartment = rs("Department")
		
    rs.close

	Call CheckBoxRead(1, "DocMaint - ReadableBy")
	
%>

<html>
<head>
    <title><%=Title%></title>

    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

    <link href="/ha_Backoffice/includes/ha_BackOfficePopup.css" type="text/css" rel="stylesheet">
    <script language="javascript" type="text/javascript" src="/includes/haFunctions.js"></script>

	<style type="text/css">
		table.filter
			{
			border:				1px solid #246B44;
			border-spacing:		2px;
			width:				85%;
			background-color:	#ECF7F1;
			margin-bottom: 		3px;
			}	
	</style>

</head>


<body onLoad="SetScreen(600,400);">
  <form method="Post" action="ha_DocumentMaintenance_post.asp?ID=<%=Request("id")%>" name="frm">

	  <span class="popupHeading"><%=Title%></span>

<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td>
	  
		<table width="100%" border="0" cellspacing="2" cellpadding="2" align="center">

			<tr><td>&nbsp;</td></tr>
			
			<tr valign="top">
				<td>
					<table width="100%" cellspacing="2" cellpadding="2" align="center">
				
			
			<tr>
				<td align="right">
					Type:
				</td>
				<td>
					<Select id="Type" name="Type">
				
<%
					strSQL = " util_HTML_option_list('ha_types_repl', 'class', 'DocMaint - Type', 'value', 'description', 'sequence', '', '" & eType & "')"
					rs.Open strSQL, connCasper10, 3, 3, 4
					If not rs.eof Then
						While not rs.eof
							Response.write Replace(rs("result"), "(blank)", "")
							rs.MoveNext
						Wend
					End If
					rs.Close
%>
		
					</Select>
				</td>		 
			</tr>

			<tr>
				<td align="right">
					Business:
				</td>
				<td>
					<Select id="Business" name="Business">
				
<%
					strSQL = " util_HTML_option_list('ha_types_repl', 'class', 'DocMaint - Business', 'value', 'description', 'sequence', '', '" & eBusiness & "')"
					rs.Open strSQL, connCasper10, 3, 3, 4
					If not rs.eof Then
						While not rs.eof
							Response.write Replace(rs("result"), "(blank)", "")
							rs.MoveNext
						Wend
					End If
					rs.Close
%>
		
					</Select>
				</td>		 
			</tr>

			<tr>
				<td align="right">
					Department:
				</td>
				<td>
					<Select id="Department" name="Department">
				
<%
					strSQL = " util_HTML_option_list('ha_types_repl', 'class', 'DocMaint - Department', 'value', 'description', 'sequence', '', '" & eDepartment & "')"
					rs.Open strSQL, connCasper10, 3, 3, 4
					If not rs.eof Then
						While not rs.eof
							Response.write Replace(rs("result"), "(blank)", "")
							rs.MoveNext
						Wend
					End If
					rs.Close
%>
		
					</Select>
				</td>		 
			</tr>
			
					</table>
				</td>
				<td align="left">
					<table cellspacing="2" cellpadding="2" class="filter">
						<%=CheckBoxFormat(1, "Readable By")%>				  
					</table>
				</td>
			
			<tr><td>&nbsp;</td></tr>
			
		</table>

    </td>
  </tr>
</table>


    <span class="popupButtons">
      <a href="javascript:document.frm.submit();" class="bo_Button80">Submit</a>
      <a href="javascript:window.close();" class="bo_Button80">Close</a>
    </span>


  </form>
</body>
</html>
