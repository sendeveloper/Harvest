  <style type="text/css">
  
	  
  </style>


<%
  
%>

<script type="text/javascript">
  
	var httpDel = new XMLHttpRequest();
	
	function clickDelete(ID, fName, row)
		{
		var eRow = 'row[' + row + ']';
		var prevColor = document.getElementById(eRow).style.backgroundColor;
		document.getElementById(eRow).style.backgroundColor="red";
		
		if(confirm('You are about to remove the record for "' + fName + '". Are you sure you want to do this?'))
			{ 
			DeletePost(ID);
			}
			
		document.getElementById(eRow).style.backgroundColor = prevColor
		}
	
	function DeletePost(ID) 
        {		
		var url = '/ha_BackOffice/includes/AJAX/ha_Delete_post.asp' +
			'?delTable=' + delTable +
			'&delID=' + ID +
			'&delUser=<%=Session("ha_shortname")%>' +
			'&Now=' + escape(Date());
		
        httpDel.open('GET', url, true);
        httpDel.onreadystatechange = DeletePostSubmit;
        httpDel.send();
        }
		
	function DeletePostSubmit()
		{
        if (httpDel.readyState == 4) 
			{
			document.frm.submit();
			}
		}

</script>