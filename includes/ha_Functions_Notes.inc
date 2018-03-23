<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<script type="text/javascript">
var Mode = 'Add';
    function clickNoteAdd(CatagoryID)
        {
        /*var URL = '<%=strPathHA%>ha_Note_Edit.asp' +
            '?ID=<%=Session("AccountID")%>' + 
            '&catID=' + CatagoryID;
        openPopUp(URL);*/
				document.getElementById('divNoteEdit').style.visibility = "visible";
				document.getElementById('spnTitle').innerText = "Add Note";
				document.getElementById('Note').value ="";
				Mode='Save';
        }
		
		

    function clickNoteEdit(NoteID)
        {
        /*var URL = '<%=strPathHA%>ha_Note_Edit.asp' +
            '?ID=<%=Session("AccountID")%>' + 
            '&nID=' + NoteID;
        openPopUp(URL);*/
		
		document.getElementById('divNoteEdit').style.visibility = "visible";
				document.getElementById('spnTitle').innerText = "Edit Note";
				document.getElementById('divNoteEdit').style.visibility = "visible";
				document.getElementById('Note').value = document.getElementById('note'+NoteID.toString()).innerText;	
					Mode='Edit';
					noteID=NoteID;
        }
</script>
<style type="text/css">
		td.head
			{
			font-weight: bold;
			color: #C0C0C0;
			border-bottom: 1px solid #C0C0C0;
			text-align: center;
			}
					
		div.divNoteEdit
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
<link href="/ha_Backoffice/includes/ha_BackOfficePopup.css" type="text/css" rel="stylesheet">
<body>	
<%
Sub NotesRead(LinkID, Category, BoxHeight)

    set rsNotes=Server.CreateObject("ADODB.Recordset")
    strSQL = "ha_Notes_read(" & LinkID & ", " & Category & ", " & BoxHeight & ", -1)"

    rsNotes.open strSQL, connCasper10, 3, 3, 4

    if not rsNotes.eof then
        do while not rsNotes.eof
            Response.write rsNotes("Result") & chr(10)
            rsNotes.movenext
        loop
    end if

    rsNotes.close
    set rsNotes = nothing

End Sub
%>
<div id="divNoteEdit" name="divNoteEdit" class="divNoteEdit">
  <form id="formEdit" name="formEdit" action="javascript:whichKey();">
	<table width="100%" cellspacing="4" cellpadding="4">
	  <tr>
	    <td width="100%" colspan="2">
		  <span id="spnTitle" name="spnTitle" style="margin: auto; 
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
	    <td  align="right">
		 Note:
		</td>
		<td>
		<textarea id="Note" name="Note"   style="height:300px;width:300px;"></textarea>
		</td>
	  </tr>
	  
	</table>
	<span id="spanCancelSave" class="popupButtons">	
			<a href="javascript:clickEditSave();" class="bo_Button80">Save</a>
			 <a href="javascript:clickEditClose();" class="bo_Button80">Close</a>
			</span>
	<span class="popupButtons">
     
    </span>
	
  </form>
  <script type="text/javascript">
  
 
 
			function clickEditClose()
			{
				document.getElementById('divNoteEdit').style.visibility = "hidden";
			}	
  </script>
</div>
</body>
</html>
