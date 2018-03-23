<%
	For i = 0 to LocationCount
	
		'Set up the first column
		If i = 0 Then
			s = "Publisher"
			resultColumns = "<td>&nbsp;</td><td>&nbsp;</td>"
		Else
			s = "Subscriber " & cStr(i)
			resultColumns = "<td id='resultInserted" & cStr(i) & "' align='right'>- - - - -</td>" & _
								"<td id='resultUpdated" & cStr(i) & "' align='right'>- - - - -</td>"
		End If
		
		'Skip a line after the Publisher row
		If i = 1 Then
			Response.write "<tr><td>&nbsp;</td><tr>"
		End If
%>
		  
      <tr>
        <td>
          <span style="color: #C0C0C0;font-size: 10px;"><%=s%></span>
        </td>
        <td>
          <%=LocationServer(i)%>
        </td>
        <td>
          <%=LocationDatabase(i)%>
<%
		If LocationFarmSource(i) = 1 Then
			Response.Write "*"
		End If
%>
        </td>
        <td>
          &nbsp;
        </td>
        <td id="resultServer<%=i%>">
          - - - - -
        </td>
        <td id="resultDatabase<%=i%>">
          - - - - -
        </td>
        <td id="resultRecords<%=i%>" align="right">
          - - - - -
        </td>
<%
	If DisplayClass = "Yes" Then
%>
        <td id="resultClasses<%=i%>" align="right">
          - - - - -
        </td>
<%
	End If
%>
        <%=resultColumns%>
<%
		If i > 0 Then
%>
		<td align="center" style="color: #C0C0C0;font-size: 10px; display: block;">
<%
			If RequestStyle = "Generic" Then
				'Generic Style
%>
		  <a href="<%=strPathServers%>TableDistribution/ha_TableDistribution_readstatus.asp
			?s=<%=LocationServer(i)%>
			&d=<%=LocationDatabase(i)%>
			&t=<%=LocationTable(i)%>
			&ExecBy=TableDistributionStatus_<%=LocationTable(i)%>.asp
			&Now=<%=now%>"
			Target="_new">Status</a>
		  <a href="<%=strPathServers%>TableDistribution/ha_TableDistribution_update.asp
			?s=<%=LocationServer(i)%>
			&d=<%=LocationDatabase(i)%>
			&t=<%=LocationTable(i)%>
			&o=<%=ObjectsClass%>
			<%=StartEndData%>
			&ExecBy=TableDistributionStatus_<%=LocationTable(i)%>.asp
			&Now=<%=now%>"
			Target="_new">Update</a>
<%
			Else
				'Old Style
%>
		  <a href="<%=strPathServers%>TableDistribution/<%=TableName%>/TableDistribution_<%=TableName%>_readstatus.asp
			?s=<%=LocationServer(i)%>
			&d=<%=LocationDatabase(i)%>
			&ExecBy=TableDistribution_<%=TableName%>.asp
			&Now=<%=now%>"
			Target="_new">Status</a>
		  <a href="<%=strPathServers%>TableDistribution/<%=TableName%>/TableDistribution_<%=TableName%>_update.asp
			?s=<%=LocationServer(i)%>
			&d=<%=LocationDatabase(i)%>
			<%=StartEndData%>
			&ExecBy=TableDistribution_<%=TableName%>.asp
			&Now=<%=now%>"
			Target="_new">Update</a>
<%
			End If
%>
		
			
		</td>
<%
		End If
%>
      </tr>
	
	
<%
	Next
%>

	</table>
