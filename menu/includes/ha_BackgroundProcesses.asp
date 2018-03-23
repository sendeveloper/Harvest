<!--#include virtual="includes/connection.asp"-->


	<table width="100%" border="0" cellspacing="5" cellpadding="5" align="left"  class="no-style">
		<tr>
			<td width="25%" valign="top">

	 <table width="100%" border="0" cellspacing="2" cellpadding="2" class="no-style">
	  
	<%
        set rs = server.createObject("ADODB.Recordset")
        SQL="ha_Menu_Dave" 
        rs.open SQL,connLocal, 3, 3, 4
        
        Response.write "<tr><td colspan='3'><h3>Background Processes</h3><hr/></td></tr>"
        
        While not rs.eof
            If rs("Sequence") < 99 Then
                If rs("CreatedBy") <> HeadingTag Then
                    HeadingTag = rs("CreatedBy")
                    Response.write "<tr><td><span style='font-weight: bold;'>" & HeadingTag & "</span></td>"
                Else
                    Response.write "<tr><td>&nbsp;</td>"			
                End If
                Response.Write "<td>" & rs("TaskName") & "</td><td>Ran " & rs("Result") & "</td></tr>"
            Else
                Select Case rs("Sequence")
                Case 100
                    ResearchStatus1 = rs("Result")
                Case 101
                    ResearchStatus2 = rs("Result")
                End Select
            End If
            rs.MoveNext
        Wend
        
        rs.close
    %>

	  </table>	  

			  <hr/>
			
			</td>
			<td width="25%" valign="top">
			
     <table width="100%" border="0" cellspacing="2" cellpadding="2" class="no-style">
	    <tr>
		  <td>
			<h3>Monthly Research Status</h3><hr/>
		  </td>
		</tr>
		
	    <tr>
		  <td>
			<%=ResearchStatus1%>
		  </td>
		</tr>
			
	    <tr>
		  <td>
			<%=ResearchStatus2%>
		  </td>
		</tr>
	  </table>

			  <hr/>
			</td>
		</tr>
	</table>
