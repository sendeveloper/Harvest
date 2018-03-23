<!--#include virtual="includes/connection.asp"-->

<html>
<head>
    <title>Work Day Description</title>

    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <link href="<%=strPathMenuIncludes%>ha_menu.css" type="text/css" rel="stylesheet">

    <script language="javascript" type = "text/javascript">

		function checkEventDate()
		{
			var d = document.getElementById('EventDate');
			if (d.value.length != 0)
            {
				if (isDate(d.value)==false)
				{
					d.focus()
					return false
				}
				return true
            }
		}

        function clickCancel()
		{
			window.close();
		}

        function clickSubmit()
		{
          //if (checkEventDate())
            //{
            document.frm.submit();
            //}
		}

    </script>

</head>

<%
   'If Request("id")="" or isnull(Request("ID")) Then
   '   response.write "<script language=javascript>"
   '  response.write "window.close()"
   '    response.write "</script>"
   ' Else
   '	Session("ID")=Request("ID")
' End If


Dim rs
'Dim SQL
set rs = server.createObject("ADODB.Recordset")

SQL = "SELECT DateWorked, Mileage, WorkDescription FROM ha_EmpHours WHERE id =  "  & Request("ID")

rs.open SQL,objcon,2,3

%>

<body style='background: #F5F9EC;'>

<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
	<tr>
		<td colspan="2">
			<table width="575px" border="0" cellspacing="0" cellpadding="0" align="center">
				<tr>
					<td>
   		              <div id="popupheader"
						<table width="100%" border="0" cellspacing="10" cellpadding="0" align="center">
							<tr>
								<th>
							    	<h3>Notes for <%=Session("ha_fname")%></h3>
									<h4><%=rs("DateWorked")%></h4>
								</th>
							</tr>
						</table>
					  </div>
					</td>
				</tr>
				<tr>
					<td>
						<table width="85%" height="335px" border="0" cellspacing="10" cellpadding="0" align="center">
<%
	if rs("Mileage") > "0" and rs("Mileage") <> ""  and IsNull(rs("Mileage")) = False then
%>
							<tr>
								<td align='center'>
									<span style='font-weight:bold; font-size:15px;'><%=rs("Mileage")%>
<%
	if rs("Mileage") < "2" then
		response.write(" Mile")
	else
		response.write(" Miles")
	end if
%>
									Recorded On This Day</span>
								</td>
							</tr>
<%
	end if
%>

							<tr>
								<td valign='top'>
									<p>
<%
	if rs("WorkDescription") <> "" or IsNull(rs("WorkDescription")) = False then

		Dim DescriptionWithReturns
		DescriptionWithReturns = rs("WorkDescription")
		DescriptionWithReturns =Replace(DescriptionWithReturns,vbcrlf,"<br>")
		Response.write DescriptionWithReturns

	end if
%>
									</p>
								</td>
							</tr>
							<tr>
								<td  style='padding-top: 30px;'align="center" colspan="2">
									<a href="javascript:window.close()" class="button">Close</a>
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
</html>
