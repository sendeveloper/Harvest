<%
	FarmSourceMessage = ""
	For i = 0 to LocationCount
		If LocationFarmSource(i) = 1 Then
			FarmSourceMessage = "* = Source Location for This Farm"
			Exit For
		End If
	Next 
%>

		<table width="100%" border="0" cellspacing="2" cellpadding="2">
		
          <tr>
		    <td>
			  <hr>
			  <%=FarmSourceMessage%>
			</td>
		  </tr>

        </table>	
