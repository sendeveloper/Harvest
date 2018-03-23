  <!--#include virtual="includes/connection.asp"-->
  
<%
    Response.buffer=true
    Response.clear

    Title = "To Do Statistics"

	Dim rs: Set rs = Server.CreateObject("ADODB.Recordset")
		
    SQL="ha_ToDo_stats"
    rs.open SQL, connLocal, 3, 3, 4

	Dim data(80,50)
	Dim t(50)
	
	If Not rs.eof Then
		employeeCount = 0
		For ii = 1 to 50
			If isnull(rs("Emp" & cStr(ii))) Then
				employeeCount = ii - 1
				Exit For
			End If
		Next
		
		Do while not rs.eof
			i = i + 1
			data(i,1) = rs("id")
			data(i,2) = rs("Class")
			data(i,3) = rs("Blank")
			data(i,4) = rs("Total")
			For ii = 1 to employeeCount
				data(i,ii+4) = rs("Emp" & cStr(ii))
			Next
			rs.MoveNext
		Loop
		recordCount = i
	End If
    rs.close
	'response.write employeeCount & "<br>"
	'employeeCount = 20
	
	colWidth = "7%"
	
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

<html>
<head>
    <title><%=Title%></title>

    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

    <link href="/ha_Backoffice/includes/ha_BackOfficePopup.css" type="text/css" rel="stylesheet">
    <script language="javascript" type="text/javascript" src="/includes/haFunctions.js"></script>

	<style type="text/css">
		td
			{
			font-size: 10pt;
			}

		td.head
			{
			font-weight: bold;
			font-size: 8pt;
			font-stretch: condensed;
			border-bottom: 1px solid black;
			}
		td.total
			{
			font-weight: bold;
			font-size: 9pt;
			border-top: 1px solid black;
			}
		h2
			{
			font-size: 12pt;
			font-weight: bold;
			border-bottom: 2px solid #C0C0C0;
			}
		span.Manager1
			{
			font-size: 8pt;
			color: #FFFFFF;
			}
		span.Manager2
			{
			font-size: 8pt;
			color: #C0C0C0;
			}
	</style
	
</head>


<body onLoad="SetScreen(1300,850);">

  <span class="popupHeading"><%=Title%></span>

<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td>
	  
		<table width="90%" border="0" cellspacing="2" cellpadding="2" align="center">

			<!----------------- Full Time ----------------------->
			<tr><td>&nbsp;</td></tr>
			<tr><td colspan="12"><h2>Full Time</h2></td></tr>

			<tr>
			  <td class="head" width="12%">
				Assignee
			  </td>
				
<%	
			colCount = 0
			For ii = 5 to employeeCount + 4
				If Data(1,ii) = "Full Time" Then
					colCount = colCount + 1
					Response.Write "<td class='head' width='" & colWidth & "' align='center'>"
					Response.Write data(2,ii)
					Response.Write "</td>"
				End If
			Next
			
			'Fill in with blank column headings
			While colCount < 11
				Response.Write "<td class='head' width='" & colWidth & "' align='center'>"
				Response.Write "&nbsp;"
				Response.Write "</td>"
				colCount = colCount + 1
			Wend
%>
			  <td width="2%">
			    &nbsp;
			  </td>
			  <td class="head" width="9%" align="right">
			    Unassigned
			  </td>
			</tr>
			
			<tr><td>&nbsp;</td></tr>
			
<%			    
	For i = 4 to 9
		colCount = 0
		Response.Write "<tr>"
			For ii = 2 to employeeCount + 4
				If Data(1,ii) = "Full Time" or ii = 2 or ii = 3 Then
					colCount = colCount + 1

					If ii <> 3 and ii <> 4 Then 'skip column 3 & 4 (Blank & Total)
						If ii = 2 Then  'col 2 (row heading)
							Response.Write "<td style='text-align: left;'>"
							If i = 4 Then
								Response.Write "<span style='color: gray;'>(blank)</span>"
							Else
								Response.Write data(i,ii)
							End If
							Response.Write "</td>"
						Else ' (row data)
							Response.Write "<td align='center'>"
							If data(i+6,ii) > 0 Then
								ManagerData1 = "<span class='Manager1'>(" & data(i + 6,ii) & ")</span>"
								ManagerData2 = "<span class='Manager2'>(" & data(i + 6,ii) & ")</span>"
								Response.Write ManagerData1 & data(i,ii) & ManagerData2
							Else
								Response.write data(i,ii)
							End If
							t(ii) = t(ii) + cInt(data(i,ii))
							Response.Write "</td>"
						End If
					End If
					
				End If
			Next
			
		'Fill in with blank columns
		While colCount < 13
			Response.Write "<td>&nbsp;</td>"
			colCount = colCount + 1
		Wend			
		
		'now do column 3
		Response.Write "<td>&nbsp;</td><td align='right'>"
		Response.Write data(i,3)
		t(3) = t(3) + cInt(data(i,3))
		Response.Write "</td>"
		Response.Write "</tr>"
	Next
%>

			<tr><td>&nbsp;</td></tr>
			
			<tr>
			  <td class="total">
				Total
			  </td>
			
<%
			colCount = 0
			For ii = 3 to employeeCount + 4
				If Data(1,ii) = "Full Time" or ii = 3 Then
					colCount = colCount + 1
					If ii <> 3 And ii <> 4 Then 'skip column 3 & 4 (blank and total)
						Response.Write "<td  class='total' align='center'>"
						Response.Write t(ii)
						Response.Write "</td>"
					End If
				End If
			Next
			
			'Fill in with blanks column headings
			While colCount < 12
				Response.Write "<td class='total'>&nbsp;</td>"
				colCount = colCount + 1
			Wend
			
			Response.Write "<td>&nbsp;</td><td class='total' align='right'>"
			Response.Write t(3)
			Response.Write "</td>"
%>
			</tr>


			<!----------------- Part Time ----------------------->
			<tr><td>&nbsp;</td></tr>
			<tr><td colspan="12"><h2>Part Time</h2></td></tr>

			<tr>
			  <td class="head" width="12%">
				Assignee
			  </td>
				
<%	
			colCount = 0
			For ii = 5 to employeeCount + 4
				If Data(1,ii) = "Part Time" Then
					colCount = colCount + 1
					Response.Write "<td class='head' width='" & colWidth & "' align='center'>"
					Response.Write data(2,ii)
					Response.Write "</td>"
				End If
			Next
			
			'Fill in with blank column headings
			While colCount < 11
				Response.Write "<td class='head' width='" & colWidth & "' align='center'>"
				Response.Write "&nbsp;"
				Response.Write "</td>"
				colCount = colCount + 1
			Wend
%>
			  <td width="2%">
			    &nbsp;
			  </td>
			  <td width="9%" align="right">
			    &nbsp;
			  </td>
			</tr>
			
			<tr><td>&nbsp;</td></tr>
			
<%			    
	For i = 4 to 9
		colCount = 0
		Response.Write "<tr>"
			For ii = 2 to employeeCount + 4
				If Data(1,ii) = "Part Time" or ii = 2 or ii = 3 Then
					colCount = colCount + 1

					If ii <> 3 and ii <> 4 Then 'skip column 3 & 4 (Blank & Total)
						If ii = 2 Then  'col 2 (row heading)
							Response.Write "<td style='text-align: left;'>"
							If i = 4 Then
								Response.Write "<span style='color: gray;'>(blank)</span>"
							Else
								Response.Write data(i,ii)
							End If
							Response.Write "</td>"
						Else ' (row data)
							Response.Write "<td align='center'>"
							If data(i+6,ii) > 0 Then
								ManagerData1 = "<span class='Manager1'>(" & data(i + 6,ii) & ")</span>"
								ManagerData2 = "<span class='Manager2'>(" & data(i + 6,ii) & ")</span>"
								Response.Write ManagerData1 & data(i,ii) & ManagerData2
							Else
								Response.write data(i,ii)
							End If
							t(ii) = t(ii) + cInt(data(i,ii))
							Response.Write "</td>"
						End If
					End If
					
				End If
			Next
			
		'Fill in with blank columns
		While colCount < 13
			Response.Write "<td>&nbsp;</td>"
			colCount = colCount + 1
		Wend			
		
		'now do column 3
		Response.Write "<td>&nbsp;</td><td align='right'>&nbsp;</td></tr>"
	Next
%>

			<tr><td>&nbsp;</td></tr>
			
			<tr>
			  <td class="total">
				Total
			  </td>
			
<%
			colCount = 0
			For ii = 3 to employeeCount + 4
				If Data(1,ii) = "Part Time" or ii = 3 Then
					colCount = colCount + 1
					If ii <> 3 And ii <> 4 Then 'skip column 3 & 4 (blank and total)
						Response.Write "<td  class='total' align='center'>"
						Response.Write t(ii)
						Response.Write "</td>"
					End If
				End If
			Next
			
			'Fill in with blanks column headings
			While colCount < 12
				Response.Write "<td class='total'>&nbsp;</td>"
				colCount = colCount + 1
			Wend
			
			Response.Write "<td>&nbsp;</td><td>&nbsp;</td>"
%>
			</tr>
			
			
			<!----------------- Contractors ----------------------->
			<tr><td>&nbsp;</td></tr>
			<tr><td colspan="12"><h2>Contractors</h2></td></tr>

			<tr>
			  <td class="head">
				Assignee
			  </td>
				
<%	
			colCount = 0
			For ii = 5 to employeeCount + 4
				If Data(1,ii) = "Contractor" Then
					colCount = colCount + 1
					Response.Write "<td class='head' align='center'>"
					Response.Write data(2,ii)
					Response.Write "</td>"
				End If
			Next
			
			'Fill in with blank column headings
			While colCount < 11
				Response.Write "<td class='head' align='center'>"
				Response.Write "&nbsp;"
				Response.Write "</td>"
				colCount = colCount + 1
			Wend
%>
			  <td>
			    &nbsp;
			  </td>
			  <td class="head" align="right">
			    Total
			  </td>			  
			</tr>
			
			<tr><td>&nbsp;</td></tr>
<%			    
	For i = 4 to 9
		colCount = 0
		Response.Write "<tr>"
			For ii = 2 to employeeCount + 4
				If Data(1,ii) = "Contractor" or ii = 2 or ii = 3 Then
					colCount = colCount + 1

					If ii <> 3 And ii <> 4 Then 'skip column 3 & 4 (blank & total)
						If ii = 2 Then  'col 2 (row heading)
							Response.Write "<td style='text-align: left;'>"
							If i = 4 Then
								Response.Write "<span style='color: gray;'>(blank)</span>"
							Else
								Response.Write data(i,ii)
							End If
							Response.Write "</td>"
						Else ' (row data)
							Response.Write "<td align='center'>"
							If data(i+6,ii) > 0 Then
								ManagerData1 = "<span class='Manager1'>(" & data(i + 6,ii) & ")</span>"
								ManagerData2 = "<span class='Manager2'>(" & data(i + 6,ii) & ")</span>"
								Response.Write ManagerData1 & data(i,ii) & ManagerData2
							Else
								Response.write data(i,ii)
							End If
							t(ii) = t(ii) + cInt(data(i,ii))
							Response.Write "</td>"
						End If
					End If
					
				End If
			Next
			
		'Fill in with blank columns
		While colCount < 13
			Response.Write "<td>&nbsp;</td>"
			colCount = colCount + 1
		Wend
			
		'now do column 4
		Response.Write "<td>&nbsp;</td><td align='right'>"
		Response.Write data(i,4)
		t(4) = t(4) + cInt(data(i,4))
		Response.Write "</td>"
		Response.Write "</tr>"
	Next
%>

			<tr><td>&nbsp;</td></tr>
			
			<tr>
			  <td class="total">
				Total
			  </td>
			
<%
			colCount = 0
			For ii = 3 to employeeCount + 4
				If Data(1,ii) = "Contractor" or ii = 3 Then
					colCount = colCount + 1
					If ii <> 3 And ii <> 4 Then 'skip column 3 & 4 (blank and total)
						Response.Write "<td class='total' align='center'>" & t(ii) & "</td>"
					End If
				End If
			Next
			
			'Fill in with blanks column headings
			While colCount < 12
				Response.Write "<td class='total'>&nbsp;</td>"
				colCount = colCount + 1
			Wend

			Response.Write "<td>&nbsp;</td><td class='total' align='right'>"
			Response.Write t(4)
			Response.Write "</td>"
%>

			</tr>
			
		</table>
		
		
		<table width="100%" border="0" cellspacing="2" cellpadding="2">
		  <tr><td>&nbsp;</td></tr>
		  <tr>
		    <td style="color: gray;" align="center">
			  Totals do not add across because you can have more than one assignee per project.
			</td>
		  </tr>
		  <tr>
		    <td style="color: gray;" align="center">
			  () Numbers = Manager Count
			</td>
		  </tr>
		  <tr><td>&nbsp;</td></tr>
		</table>
    </td>
  </tr>

  <tr>
    <td>
      <span class="popupButtons">
		<a href="javascript:window.close();" class="bo_Button80">Close</a>
      </span>
	</td>
  </tr>
	
  <tr><td>&nbsp;</td></tr>
  
</table>

</body>
</html>
