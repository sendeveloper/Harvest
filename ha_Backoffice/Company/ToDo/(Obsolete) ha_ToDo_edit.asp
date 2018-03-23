<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/Config.asp"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsHelp.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->
  <!--#Include virtual="ha_backoffice/includes/FunctionsNotes.inc"-->
  <!--#include virtual="ha_backoffice/includes/Dates/ConvDate.inc"-->
  
<%
    Dim rs, rsSec, strSQL, ID
    Set rs = Server.CreateObject("ADODB.Recordset")

	ID 				= 0
	ePriority 		= ""
	eTitle 			= ""
	eDescription 	= ""
	eAssignedTo 	= ""
	eManagedBy		= ""
	eServer 		= ""
	eCategory 		= ""
	eStatus			= ""
	eCreatedBy		= ""
	eCreatedDate	= ""
	eEditedBy		= ""
	eEditedDate		= ""
	eNotes			= ""

    If Request("ID") = "" Then
		Title = "Add To Do"
	Else
		Title = "Edit To Do"
		strSQL = "SELECT * FROM ha_ToDo WHERE ID = " & Request("ID")
		rs.Open strSQL, connLocal, 1, 3
		
		If Not rs.eof Then
			ID 				= Null2Space(rs("ID"))
			ePriority 		= Null2Space(rs("Priority"))
			eTitle 			= Null2Space(rs("Title"))
			eDescription 	= Null2Space(rs("Description"))
			eAssignedTo 	= Null2Space(rs("AssignedTo"))
			eManagedBy	 	= Null2Space(rs("ManagedBy"))
			eServer 		= Null2Space(rs("Server"))
			eCategory 		= Null2Space(rs("Category"))
			eStatus			= Null2Space(rs("Status"))
			eCreatedBy		= Null2Space(rs("CreatedBy"))
			eCreatedDate	= Null2Space(rs("CreatedDate"))
			eEditedBy		= Null2Space(rs("EditedBy"))
			eEditedDate		= Null2Space(rs("EditedDate"))
			eNotes			= Null2Space(rs("Notes"))
		End If
		
		rs.Close	
    End If


%>

<%



Set objCmd  = Server.CreateObject("ADODB.Command")
Set rs = Server.CreateObject("ADODB.Recordset")
    
    If Request.QueryString("a") = "Save" Then
      strSQL = "SELECT TOP 1 * FROM ha_Notes"
      rs.Open strSQL, connLocal, 2, 3
      rs.AddNew
	    rs("CreatedBy") = Session("ha_shortname")
		rs("CreatedDate") = Now
		
    Elseif Request.QueryString("a") = "Edit" then
      strSQL = "SELECT * FROM ha_Notes WHERE ID = " & Request("noteid")
	  'response.write strSQL
	  'response.end
      rs.Open strSQL, connLocal, 2, 3
		
    End If
	
	if request.querystring("a") = "Save" or Request.QueryString("a") = "Edit" then
		rs("LinkID") = Request("ID")
		rs("Category") = "ToDo"
		
		rs("Note") = replace(replace(ltrim(rtrim(Request("Note"))),"~~","&"),"^^","#")
		rs("EditedBy") = Session("ha_shortname")
		rs("EditedDate") = Now
		rs.Update
		rs.Close
			response.redirect "ha_ToDo_edit.asp?ID=" & Request("ID")
	end if




	
	
%>
<!DOCTYPE html>     
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
	<title>Harvest American Backoffice - <%=Title%></title>
  
	<link href="/ha_BackOffice/includes/ha_BackofficePopup.css" type="text/css" rel="stylesheet" />
	<script type="text/javaScript" src="/ha_BackOffice/includes/haFunctions.js"></script>
	
	<script type="text/javascript">
      var e = '<ul>';
      var fldcount = 2;
      var notval = false;
		var noteID = 0;
      function Validate(field) {
          var f = document.getElementById(field).value;

          if (f != '' && f.length > 0) { fldcount--; }
          else { e = e + '<li>Field ' + field.substring(3) + ' is required.</li>\n'; }

          if (fldcount == 0) { notval = true; }
      }

      function Submit() {
          Validate('Description');
          Validate('Category');

          if (notval == true) { document.frm.submit(); }
          else {
              fldcount = 2;
              e = e + '</ul>';
              var d = document.getElementById('errorcon');
              d.innerHTML = e;
          }
      }
      
      function clickEditSave()
			{
			//document.getElementById('divNoteEdit').style.visibility = "hidden";
			var note = document.getElementById('Note').value;
		
			var URL = 'ha_ToDo_edit.asp?Mode='+ Mode +'&a='+ Mode +'&ID=<%=ID%>&noteid='+ noteID +'&Category=ToDo&Note='+note.replace('&','~~').replace('#','^^');
			window.location.href = URL;
			}	
			
	</script>
	
  <style type="text/css">
			
    input[type="text"]
		{
		font-family: Verdana, Arial, Helvetica, sans-serif;
		width: 40em;   
		}
	
    select
		{
		font-family: Verdana, Arial, Helvetica, sans-serif;
		width: 12.5em;   
		}
	
	textarea
		{
		font-family: Verdana, Arial, Helvetica, sans-serif;
		}
	
    textarea.Description
		{
		width: 40em;   
		height: 12em; 
		resize: none; 
		}
		
    textarea.Notes
		{
		width: 40em;   
		height: 8em; 
		resize: none; 
		}
			
    #errorcon
		{
        font-size: .75em;
        color: Red;
        font-weight: bold;
		}
		
	tr
		{
		vertical-align: top;
		}
	
  </style>
</head>

<body onload="SetScreen(1200,950)">
  <form action="ha_ToDo_post.asp?ID=<%=ID%>" method="post" name="frm">
<inpu type="hidden" name="editurl" id="editurl"/>
	<span class="popupHeading"><%=Title%></span>
	
	<table width="95%" border="0" cellspacing="2" cellpadding="2" style="margin: 0 auto;">
	  <tr>
	    <td width="70%" vAlign="top">
	
		<table width="100%" border="0" cellspacing="2" cellpadding="2">

			<tr>
			  <td width="15%">&nbsp;</td>
			  <td width="70%">&nbsp;</td>
			  <td width="15%">&nbsp;</td>
			</tr>
  
<%
		If Request("tMode") = "Edit" Then
%>
			<tr>
				<td align="right">
					ID:
				</td>
				<td>
					<%=ID%>
				</td>
			</tr>

<%
		End If
%>
  
  
			<tr>
			  <td align="right">
				Priority:
			  </td>
			  <td>
				<Select id="Priority" name="Priority">
  
<%
					strSQL = " util_HTML_option_list('ha_types_repl', 'class', 'ToDo - PriorityValues', 'value', 'description', 'sequence', '', '" & ePriority & "')"
					rs.Open strSQL, connLocal, 3, 3, 4
					If not rs.eof Then
						While not rs.eof
							Response.write Replace(rs("result"), "(blank)", "")
							rs.MoveNext
						Wend
					End If
					rs.Close
%>
		
				</Select>
				<%=PopupHelp(10)%>
			  </td>
			</tr>
  		
			<tr vAlign="top">
			  <td align="right">
				Title:
			  </td>
			  <td>
				<input type="text" id="Title" name="Title" size="100" value="<%=Replace(eTitle,chr(34),"&quot;")%>">
			  </td>
			  <td>
				<%=PopupHelp(32)%>
			  </td>
			</tr>
			
			<tr>
			  <td align="right">
				Description:
			  </td>
			  <td>
				<textarea cols="1" rows="1" id="Description" name="Description" class="Description"><%=eDescription%></textarea>
			  </td>
			  <td>
				<%=PopupHelp(11)%>
			  </td>
			</tr>
			
			<tr vAlign="top">
			  <td align="right">
				Server Name:
			  </td>
			  <td>
				<Select id="Server" name="Server">
  
<%
					strSQL = " util_HTML_option_list('ha_types_repl', 'class', 'Domains - Server', 'value', 'description', 'sequence', '', '" & eServer & "')"
					rs.Open strSQL, connLocal, 3, 3, 4
					If not rs.eof Then
						Response.write "<OPTION value=''></OPTION>"
						While not rs.eof
							Response.write Replace(rs("result"), "(blank)", "")
							rs.MoveNext
						Wend
					End If
					rs.Close
%>
		
				</Select>
			  </td>
			</tr>
							
						
			<tr>
			  <td align="right">
				Category:
			  </td>
			  <td>
				<Select id="Category" name="Category">
  
<%
					strSQL = " util_HTML_option_list('ha_types_repl', 'class', 'ToDo - Category', 'value', 'description', 'sequence', '', '" & eCategory & "')"
					rs.Open strSQL, connLocal, 3, 3, 4
					If not rs.eof Then
						While not rs.eof
							Response.write Replace(rs("result"), "(blank)", "")
							rs.MoveNext
						Wend
					End If
					rs.Close
%>
		
				</Select>
				<%=PopupHelp(13)%>
			  </td>
			</tr>
			  
			<tr>
			  <td align="right">
				Status:
			  </td>
			  <td>
				<Select id="Status" name="Status">
  
<%
					strSQL = " util_HTML_option_list('ha_types_repl', 'class', 'ToDo - Status', 'value', 'value', 'sequence', '', '" & eStatus & "')"
					rs.Open strSQL, connLocal, 3, 3, 4
					If not rs.eof Then
						While not rs.eof
							Response.write Replace(rs("result"), "(blank)", "")
							rs.MoveNext
						Wend
					End If
					rs.Close
%>
		
				</Select>
				<%=PopupHelp(14)%>
			  </td>
			</tr>
													
			<tr vAlign="top">
			  <td align="right">
				Notes:
			  </td>
			  <td>
				<textarea cols="1" rows="1" id="Notes" name="Notes" class="Notes"><%=eNotes%></textarea>
			  </td>
			  <td>
				<%=PopupHelp(15)%>
			  </td>
			</tr>
			  			  
		</table>
		
		</td>
	    <td width="30%" vAlign="top">
  
			<table width="100%" border="0" cellspacing="2" cellpadding="2">
			  <tr>
			    <td vAlign="top">
				  <br><br>
				  <span style="font-weight: bold; border-bottom: 1px solid black;">Assigned To</span><br><br>
<%
					strSQL = "ha_ToDo_AssignedTo_list(" & Request("ID") & ")"
					rs.Open strSQL, connLocal, 3, 3, 4
					If not rs.eof Then
						While not rs.eof
							Response.write rs("HTMLAssignedTo") & "<br>" & vbCrLf
							rs.MoveNext
						Wend
					End If
					rs.Close
%>				  
				</td>
			    <td vAlign="top">
				  <br><br>
				  <span style="font-weight: bold; border-bottom: 1px solid black;">Managed By</span><br><br>
<%
					strSQL = "ha_ToDo_AssignedTo_list(" & Request("ID") & ")"
					rs.Open strSQL, connLocal, 3, 3, 4
					If not rs.eof Then
						While not rs.eof
							Response.write rs("HTMLManagedBy") & "<br>" & vbCrLf
							rs.MoveNext
						Wend
					End If
					rs.Close
%>				  
				</td>
			  </tr>
			</table>
  
		</td>
	  </tr>
	  
	  <!---------- Notes ---------->
	  <tr>
		<td colspan="2">
		  <table width="100%" cellspacing="0" cellpadding="0">
		  
			<tr>
			  <td align="left">
				<b>To Do Notes</b>
			  </td>
			  <td align="right">
				<a href="javascript:clickNoteAdd('ToDo');" 
				  class="headingLink">Add New To Do Note</a>
			  </td>
			</tr>
				  
			<tr>
			  <td colspan="2">
				<table width="100%" cellspacing="0" cellpadding="0">
				  
				  <tr>
					<td colspan="2">
					  <table width='100%' cellspacing='0' cellpadding='0'>
					  
						<tr valign="top">
						  <td width="100%">
							<%Call NotesRead(ID, "ToDo", 230)%>
						  </td>
						</tr>
						
					  </table>
					</td>
				  </tr>
				  
				</table>
			  </td>
			</tr>

		  </table>
		</td>
	  </tr>
	</table>
  
	
    <span class="popupCredit"><%=CreditLine(eCreatedBy, eCreatedDate, eEditedBy, eEditedDate)%></span>	

    <span class="popupButtons">
      <a href="javascript:document.frm.submit();" class="bo_Button80">Submit</a>
      <a href="javascript:window.close();" class="bo_Button80">Cancel</a>
    </span>
  
	<div id="errorcon">&nbsp;</div>
  
  </form>
  
    
</body>
</html>
