<style type="text/css">

	td.PageHeadingHead
		{
		font-size: 		18px;
		font-weight: 	bold;
		color: 			#C9BB9E;
		padding-left:	.25em;
		padding-right:	.25em;
		}
		
</style>


  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
	  <td>
	    <table width="100%" border="0" cellspacing="0" cellpadding="0">
		  <tr>
			<td class="PageHeadingHead"> 
			  Harvest American - Back Office
			</td>
			
			<td align="right" class="PageHeadingHead">
			  <%=Session("ha_fname") & " " & Session("ha_lname")%>
			  <span style="font-size: 8pt;">
				<a href="/includes/logout.asp" style="color: #C9BB9E;">LOGOUT</a>
			  </span>
			</td>
		  </tr>
		</table>
      </td>
	</tr>
	
	
    <tr> 
      <td align="center" style="padding-right: 2px; 
								font-size: 24px;
								font-weight: bold;
								color: #000000;"> 
        <%=PageHeading%>
      </td>
	</tr>

	<!--#include virtual="ha_Backoffice/includes/menu/PageMenu.inc"-->
	
  </table>
