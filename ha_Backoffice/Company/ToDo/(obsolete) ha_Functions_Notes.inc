<link href="/ha_Backoffice/includes/ha_BackOfficePopup.css" type="text/css" rel="stylesheet">
<script language="javascript" type="text/javascript" src="/includes/haFunctions.js"></script>
<style type="text/css">
		td.head
			{
			font-weight: bold;
			color: #C0C0C0;
			border-bottom: 1px solid #C0C0C0;
			text-align: center;
			}
					
		div.divEdit
			{
			position:absolute; 
			border: 2px solid black;
	        top: 100px; 
        	left: 200px; 
	        width: 450px; 
        	height: 450px;
	        background-color: #FFFFCC;
        	visibility: hidden;
			}
			
		select 
			{ 
			height: 20px;
			}
	</style>
	
<div id="divNoteEdit" name="divNoteEdit" class="divNoteEdit">
  <form id="formEdit" name="formEdit" action="javascript:whichKey();">
	<table width="100%" cellspacing="4" cellpadding="4">
	  <tr>
	    <td width="100%" colspan="2">
		  <span style="margin: auto; 
			font-weight: bold; 
			font-size: 16pt; 
			display: block;
			text-align: center;
			border-bottom: 1px solid black;">
			Note Edit
		  </span>
		</td>
	  </tr>
	  
	  <tr>
		<td style="width: 30%; text-align: right;">
		  LinkID:
		</td>
		<td style="width: 70%; font-weight: bold;">
		  <select id="linkid" name="linkid">
		  <option value="597">Dummy1</option>
		  <option value="598">Dummy2</option>
		  <option value="599">Dummy3</option>
		  </select>
		</td>
	  <tr>
	  
	  <tr>
		<td style="text-align: right;">
		 Category:
		</td>
		<td style="font-weight: bold;">
		  <select id="category" name="category">
		  <option value="597">Dummy1</option>
		  <option value="598">Dummy2</option>
		  <option value="599">Dummy3</option>
		  </select>
		</td>
	  </tr>
	    
	  <tr>
	    <td  align="right">
		 Note:
		</td>
		<td>
		<textarea id="note" name="note"   style="height:100px;width:300px;"></textarea>
		</td>
	  </tr>
	  
	</table>
	<span id="spanCancelSave" class="popupButtons">	
			<a href="javascript:clickEditSave();" class="bo_Button80">Save</a>
			 <a href="javascript:clickAddClose();" class="bo_Button80">Close</a>
			</span>
	<span class="popupButtons">
     
    </span>
	
  </form>
  <script type="text/javascript">
  
 
  function clickEditSave()
			{
			document.getElementById('divEdit').style.visibility = "hidden";
			
			var URL = 'ha_Note_edit.asp?a=Save&ID=<%=ID%>&LinkID=<%=Request("LinkID")%>&Category=<%=Request("Category")%>&Note=<%=Request("Note")%>';
			//alert(URL);
			window.location.href = URL;
			}	
			
			function clickEditClose()
			{
			document.getElementById('divNoteEdit').style.visibility = "hidden";
			}	
  </script>
</div>