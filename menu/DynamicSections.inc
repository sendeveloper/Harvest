<script language='JavaScript'>

    var httpDynamicMenu = new XMLHttpRequest();
    var httpBackgroundProcesses = new XMLHttpRequest();	
	
	function runStaffStatus() 
        {		
		var url = '/menu/includes/ha_StaffStatus.asp';
		
        httpDynamicMenu.open('GET', url, true);
        httpDynamicMenu.onreadystatechange = getStaffStatusResponse;
        httpDynamicMenu.send();
        }

    function getStaffStatusResponse() 
        {
        if (httpDynamicMenu.readyState == 4) 
			{
			if (httpDynamicMenu.status == 200) 
				{
				document.getElementById('StaffStatusBox').innerHTML = httpDynamicMenu.responseText;
				}
			}
        }
				
	function runBackgroundProcesses() 
        {		
		var url = '/menu/includes/ha_BackgroundProcesses.asp';
		
        httpBackgroundProcesses.open('GET', url, true);
        httpBackgroundProcesses.onreadystatechange = getBackgroundProcessesResponse;
        httpBackgroundProcesses.send();
        }

    function getBackgroundProcessesResponse() 
        {
        if (httpBackgroundProcesses.readyState == 4) 
			{
			if (httpBackgroundProcesses.status == 200) 
				{
				document.getElementById('BackgroundProcessesBox').innerHTML = httpBackgroundProcesses.responseText;
				}
			}
        }
				
	runStaffStatus();
	runBackgroundProcesses();

</script>

<table width="100%" border="0" cellspacing="5" cellpadding="5" align="left" class="no-style">
  <tr>

    <td width="50%" valign="top" id="StaffStatusBox">
		  <img src="../images/AjaxLoading.gif" style="display: block; margin-left: auto; margin-right: auto;">
        <hr/>
    </td>

    <td width="50%" valign="top" id="BackgroundProcessesBox">
	<table width="100%" border="0" cellspacing="5" cellpadding="5" align="left" class="no-style">
		<tr>
			<td width="25%" valign="top">
			  <hr/>
					  <img src="../images/AjaxLoading.gif" style="display: block; margin-left: auto; margin-right: auto;">
			  <hr/>
			
			</td>
			<td width="25%" valign="top">
			
			  <hr/>
					  <img src="../images/AjaxLoading.gif" style="display: block; margin-left: auto; margin-right: auto;">
			  <hr/>
			</td>
		</tr>
	</table>
 

    </td>
  </tr>

</table>
