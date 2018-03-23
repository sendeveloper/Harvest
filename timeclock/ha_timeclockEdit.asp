<!doctype html>
<!--#include virtual="includes/connection.asp"-->

<html>
<head>
  <title>Time Card Record | 
<%
    Dim strPathBOInc : strPathBOInc = strPathBase & "/ha_Backoffice/Includes/"
    If Session("cMode") = "Edit" and IsNull(eDatePaid) = False Then
	  response.write("Locked")
	Else
%>
	<%=Session("cMode")%>
<%
    End If
%>
  </title>

  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
  <link href="<%=strPathMenuIncludes%>ha_menu.css" type="text/css" rel="stylesheet" />
  <link href="<%=strPathBOInc%>datepick/jquery-ui.css" rel="stylesheet" />
  <script type="text/javascript">
      window.resizeTo(775, 575);
  </script>
  <%
	If Request("cMode")="Add" Then
	  Session("cMode")="Add"
	Else
	  Session("cMode")="Edit"
	End If
  %>

</head>

<%
    'If Request("id")="" or isnull(Request("ID")) Then
    '   response.write "<script language=javascript>"
    '  response.write "window.close()"
    '    response.write "</script>"
    ' Else
    	Session("ID")=Request("ID")
   ' End If


    Dim rs ', SQL
    set rs = server.createObject("ADODB.Recordset")

   If Session("cMode")="Add" then
	 Title = "Add New Day"
	 If Not Request("empid") = "" Then eEmpID = Request("empid") Else eEmpID = Session("ha_EmpID")
	 eDateWorked = date()
	 eQuantityHours = "8"
	 eHoursType = 1
	 eDatePaid = ""
	 eWorkDescription = ""
	 eMileage = ""
   Else
     SQL="SELECT * FROM ha_EmpHours WHERE [ID] = " & Request("ID")
     'response.write SQL & chr(13)
     rs.open SQL, connCasper10, 2, 3
     eID = rs("ID")
     eEmpID = rs("EmpID")
     eDateWorked = rs("DateWorked")
     eQuantityHours = rs("QuantityHours")
     eHoursType = rs("HoursType")
     eDatePaid = rs("DatePaid")
     eWorkDescription = rs("WorkDescription")
     eMileage = rs("Mileage")
     rs.Close
   End If
%>


<body style='background: #F5F9EC;'>
  <form method="post" action="ha_timeclockPost.asp" name="frm">
    <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
	  <tr>
		<td colspan="2">
		  <table width="575px" border="0" cellspacing="0" cellpadding="0" align="center">
			<tr>
			  <td>
				<div id="popupheader">
				  <table width="100%" border="0" cellspacing="10" cellpadding="0" align="center">
					<tr>
					  <th>
					    <h3>Time Card Record</h3>
                        <%
                        If Session("cMode") = "Edit" and IsNull(eDatePaid) = False Then
                        %>
						<h4 style="color:red;">Locked</h4>
                        <%
	                    Else
                        %>
						<h4><%=Session("cMode")%></h4>
                        <%
                        End If
                        %>
					  </th>
					</tr>
				  </table>
				</div>
				<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
				  <tr>
					<td>
					  <table style='background: #CDE2A1;' width="91%" height='230px' border="0" cellspacing="0" cellpadding="3" align="right">
						<tr style="display:none">
						  <td>Employee ID:
							<input type="text" name="EmpID" value="<%=eEmpID%>" size="3" readonly="readonly" />
						  </td>
						</tr>
						<tr>
						  <td align="center" valign="bottom">
							<span style='font-weight:bold; font-size:15px;'>Date Worked</span>
						  </td>
						</tr>
						<tr>
						  <td align='center' valign="top">
                            <%
                            If Session("cMode") = "Edit" and IsNull(eDatePaid) = False Then
                            %>
							<input type="date" style="text-align: center;" name="DateWorked" class="datepick" value="<%Response.write Year(eDateWorked) & "-"
                            If Month(eDateWorked) > 10 Then
                            Response.write "0" & Month(eDateWorked) & "-"
                            Else
                            Response.write Month(eDateWorked) & "-"
                            End If 
                            
                            If Day(eDateWorked) > 10 Then
                            Response.write "0" & Day(eDateWorked) & "-"
                            Else
                            Response.write Day(eDateWorked) & "-"
                            End If 
                            Day(eDateWorked)%>" size="10" />
                            <%
	                        Else
                            %>
							<input type="date" style="text-align: right;" name="DateWorked" class="datepick" value="<%
                            Response.write Year(eDateWorked) & "-"
                            If Month(eDateWorked) < 10 Then
                            Response.write "0" & Month(eDateWorked) & "-"
                            Else
                            Response.write Month(eDateWorked) & "-"
                            End If 
                            
                            If Day(eDateWorked) < 10 Then
                            Response.write "0" & Day(eDateWorked)
                            Else
                            Response.write Day(eDateWorked)
                            End If 
                            Day(eDateWorked)%>" size="10" />
							<!--<a href="#" onclick="cdp1.showCalendar(this, 'DateWorked'); return false;" title="View Calendar"><img style="padding-left: 5px; " align="absmiddle" src="<%=strPathMenuImages%>/calendar.png" border="0"></a>-->
                            <%
                            End If
                            %>
                            <script type="text/javascript" src="<%=strPathBOInc%>datepick/jquery-1.9.1.js"></script>
                            <script type="text/javascript" src="<%=strPathBOInc%>datepick/jquery-ui.js"></script>
                            <script type="text/javascript">
                            (function () {
                                var elem = document.createElement('input');
                                elem.setAttribute('type', 'date');

                                if (elem.type === 'text') {
                                    $('.datepick').datepicker({
                                        dateFormat: 'yy-mm-dd',
                                        changeMonth: true,
                                        changeYear: true,
                                        showOn: 'button',
                                        buttonImage: '/menu/images/calendar.png',
                                        buttonImageOnly: true
                                    });

                                    $('.datepick').attr('readonly', 'readonly');
                                }
                            })(); 
                            </script>

						  </td>
						</tr>
						<tr>
						  <td align="center" valign="bottom">
							<span style='font-weight:bold; font-size:15px;'>Hours Worked</span>
						  </td>
						</tr>
						<tr>
						  <td align='center' valign="top">
                            <%
                            If Session("cMode") = "Edit" and IsNull(eDatePaid) = False Then
                            %>
							<input type="text" style="text-align: center;" name="QuantityHours" value="<%=eQuantityHours%>" size="2" readonly="readonly" />
                            <%
	                        Else
                            %>
							<input type="text" style="text-align: right;" name="QuantityHours" value="<%=eQuantityHours%>" size="2" />
                            <%
                            End If
                            %>
						  </td>
						</tr>
						<tr>
						  <td align="center" valign="bottom">
							<span style='font-weight:bold; font-size:15px;'>Hours Type</span>
						  </td>
						</tr>
						<tr>
						  <td align='center' valign="top">
                            <%
                            'If (Session("cMode") = "Edit" and IsNull(eDatePaid) = False) or (Session("Vacation") = False and Session("PaidTimeOff") = False And Not Session("ha_Permissions") = "99") Then
                            %>
                            <!--<select name="HoursType">
                              <option value="1">Regular</option>
                              <option value="4">Holiday</option>
                              <% 'If Session("Vacation") = True Then %>
                              <option value="2">Vacation</option>
                              <% 'End If 
                              'If Session("PaidTimeOff") = True Then %>
                              <option value="3">Personal</option>
                              <% 'End If %>
                            </select>
                            <input type="text" name="HoursType" value="<%=eHoursType%>" size="12" style="display:none" readonly="readonly" />-->
                            <%
	                        'Else
                            %>
							<select name="HoursType">
                            <%
                            strSQL = "util_HTML_option_list('Casper10.ha_backoffice.dbo.ha_types_repl', 'class', 'HoursType', 'value', 'description', 'sequence', '', '" & eHoursType & "')"

                            rs.open strSQL, connCasper10, 3, 3, 4

                            If not rs.eof Then
                              While not rs.eof
                              Response.Write rs("Result")
                              rs.movenext
                              Wend
                            End If

                            rs.close
                            %>
                            </select>
                            <%
                            'End If
                            %>
                          </td>
						</tr>
					  </table>
					</td>
					<td>
					  <table width="85%" border="0" cellspacing="0" cellpadding="0" align="center">
						<tr>
						  <td>
							<span style='font-weight:bold; font-size:15px;'>Notes</span>
						  </td>
						</tr>
						<tr>
						  <td align='center' valign="top">
                            <%
                            If Session("cMode") = "Edit" and IsNull(eDatePaid) = False Then
                            %>
							<textarea name="WorkDescription" cols="40" rows="10" readonly="readonly"><%=eWorkDescription%></textarea>
                            <%
	                        Else
                            %>
							<textarea name="WorkDescription" cols="40" rows="10"><%=eWorkDescription%></textarea>
                            <%
                            End If
                            %>
						  </td>
						</tr>
					  </table>
					</td>
				  </tr>
				  <tr>
					<td colspan='2'>
					  <table style='background: #CDE2A1;' width="548px" height='30px' border="0" cellspacing="0" cellpadding="3" align="center">
						<tr>
						  <td align='center'>
                            <%
                            If Session("cMode") = "Edit" and IsNull(eDatePaid) = False Then
                            %>
							<span style='font-weight:bold; font-size:15px;'>Mileage: </span><input type="text" style="text-align: center" name="Mileage" value="<%=eMileage%>" size="2" readonly="readonly">
                            <%
	                        Else
                            %>
							<span style='font-weight:bold; font-size:15px;'>Mileage: </span><input type="text" style="text-align: right" name="Mileage" value="<%=eMileage%>" size="2">
                            <%
                            End If
                            %>
						  </td>
						</tr>
					  </table>
					</td>
				  </tr>
				</table>
			  </td>
			</tr>
            <%
            If Session("cMode") = "Edit" and IsNull(eDatePaid) = False Then
            %>
			<tr>
			  <td style='padding-top: 10px;' align="center" colspan="2">
				<span style='font-weight:bold; font-size:10px; color: red;'>This record is locked. Payment for this record was issued on <%=eDatePaid%>.</span>
			  </td>
			</tr>
			<tr>
			  <td style='padding-top: 30px;' align="center" colspan="2">
				<a href="javascript:window.close()" class="button">Close</a>
			  </td>
			</tr>
            <%
	        Else
            %>
			<tr>
			  <td style='padding-top: 50px;' align="center" colspan="2">
				<a href="javascript:document.frm.submit()" class="button">Submit</a>
				<a href="javascript:window.close()" class="button">Cancel</a>
			  </td>
			</tr>
            <%
            End If
            %>
		  </table>
		</td>
	  </tr>
    </table>
  </form>
</body>
</html>