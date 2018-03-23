		<table width="100%" border="0" cellspacing="2" cellpadding="2">
		
          <tr>
            <td class="popupHeading" style="width: 100%;">
              Table Distribution Status - <%=TableName%><%=UpdateClass%>
            </td>
          </tr>
		  
		</table>
		
		  
		<table width="100%" border="0" cellspacing="2" cellpadding="2">
		
          <tr>
		    <td class="databaseHead" colspan="3">
			  Locations
			</td>
			<td>
			  &nbsp;
			</td>

<%
	If DisplayClass = "Yes" Then
%>
		    <td class="databaseHead" colspan="7">
<%
	Else
%>
		    <td class="databaseHead" colspan="6">
<%
	End If
%>
			  Results
			</td>
		  </tr>
		  
		  <tr>
			<td class="subHead2" colspan="6">
			  &nbsp;
			</td>
			<td class="subHead2">
			  Record
			</td>
			
<%
	If DisplayClass = "Yes" Then
%>
        <td class="subHead2">
          Class
        </td>
<%
	End If
%>
		
			<td class="subHead2">
			  Inserted
			</td>
			<td class="subHead2">
			  Updated
			</td>
			<td class="subHead2">
			  Debug
			</td>
		  </tr>
		  
		  
		  <tr>
			<td width="10%">
			  &nbsp;
			</td>
			<td class="subHead" width="10%">
			  Server
			</td>
			<td class="subHead" width="15%">
			  Database
			</td>
			<td width="2%">
			  &nbsp;
			</td>
			<td class="subHead" width="10%">
			  Server
			</td>
			<td class="subHead" width="15%">
			  Database
			</td>
			<td class="subHead" width="7%">
			  Count
			</td>
			
<%
	If DisplayClass = "Yes" Then
%>
			<td class="subHead" width="7%">
			  Count
			</td>
<%
	End If
%>			
			
			<td class="subHead" width="7%">
			  Count
			</td>
			<td class="subHead" width="7%">
			  Count
			</td>
			<td class="subHead" width="12%">
			  AJAX URLs
			</td>
		  </tr>
