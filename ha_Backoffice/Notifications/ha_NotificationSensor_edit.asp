<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/Config.asp"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsHelp.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->
  <!--#Include virtual="ha_backoffice/includes/FunctionsNotes.inc"-->
  
<%
    Dim rs, rsSec, strSQL, ID
    Set rs = Server.CreateObject("ADODB.Recordset")

	ID 				= 0
	eTitle 			= ""
	eDescription 	= ""
	eCategory 		= ""
	eServer			= ""
	eIncrement 		= ""
	eStatus			= ""
	eCreatedBy		= ""
	eCreatedDate	= ""
	eEditedBy		= ""
	eEditedDate		= ""
	eNotes			= ""

    If Request("ID") = "" Then
		Title = "Add Notification Sensor"
	Else
		Title = "Edit Notification Sensor"
		strSQL = "ha_NotificationSensor_read(" & Request("ID") & ")"
		rs.Open strSQL, connLocal, 3, 3, 4
		
		If Not rs.eof Then
			ID 				= Null2Space(rs("ID"))
			eTitle 			= Null2Space(rs("Title"))
			eDescription 	= Null2Space(rs("Description"))
			eCategory 		= Null2Space(rs("Category"))
			eServer			= Null2Space(rs("ServerName"))
			eIncrement		= Null2Space(rs("Increment"))
			eStatus			= Null2Space(rs("Status"))
			eCreatedBy		= Null2Space(rs("CreatedBy"))
			eCreatedDate	= Null2Space(rs("CreatedDate"))
			eEditedBy		= Null2Space(rs("EditedBy"))
			eEditedDate		= Null2Space(rs("EditedDate"))
		End If
		
		rs.Close	
    End If


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

<body onload="SetScreen(800,650)">
  <form action="ha_NotificationSensor_post.asp?ID=<%=ID%>" method="post" name="frm">

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
    		
			<tr vAlign="top">
			  <td align="right">
				Title:
			  </td>
			  <td>
				<input type="text" id="Title" name="Title" size="100" value="<%=Replace(eTitle,chr(34),"&quot;")%>">
			  </td>
			  <td>
				<%=PopupHelp(38)%>
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
				<%=PopupHelp(39)%>
			  </td>
			</tr>
							
						
			<tr>
			  <td align="right">
				Category:
			  </td>
			  <td>
				<Select id="Category" name="Category">
  
<%
					strSQL = " util_HTML_option_list('ha_types_repl', 'class', 'Notification Sensor Categories', 'value', 'description', 'sequence', '', '" & eCategory & "')"
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
				<%=PopupHelp(40)%>
			  </td>
			</tr>
			
			  
			<tr>
			  <td align="right">
				Server:
			  </td>
			  <td>
				<Select id="ServerName" name="ServerName">
  
<%
					strSQL = "ha_BackOffice.dbo.util_HTML_option_list('ha_PublishedTables.dbo.ha_types', 'class', 'Domains - Server', 'description', 'description', 'sequence', '', '" & eServer & "')"
		
				rs.Open strSQL, connLocal, 3, 3, 4
				If not rs.eof Then
					While not rs.eof
						Response.write rs("result")
						rs.MoveNext
					Wend
				End If
				rs.Close
%>
		
				</Select>
				<%=PopupHelp(42)%>
			  </td>
			</tr>
			
			
			<tr>
			  <td align="right">
				Increment:
			  </td>
			  <td>
				<Select id="Increment" name="Increment">
  
<%
					strSQL = " util_HTML_option_list('ha_types_repl', 'class', 'Notification Increment', 'value', 'description', 'sequence', '', '" & eIncrement & "')"
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
				<%=PopupHelp(43)%>
			  </td>
			</tr>
			
			<tr>
			  <td align="right">
				Status:
			  </td>
			  <td>
				<Select id="Status" name="Status">
  
<%
					strSQL = " util_HTML_option_list('ha_types_repl', 'class', 'Notification Sensor Status', 'value', 'value', 'sequence', '', '" & eStatus & "')"
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
				<%=PopupHelp(41)%>
			  </td>
			</tr>
			
			<tr><td>&nbsp;</td></tr>
			<tr><td>&nbsp;</td></tr>
			<tr><td>&nbsp;</td></tr>
			<tr><td>&nbsp;</td></tr>
			
		</table>
		
		</td>
	  </tr>  
	</table>
  	
    <span class="popupCredit"><%=CreditLine(eCreatedBy, eCreatedDate, eEditedBy, eEditedDate)%></span>	

    <span class="popupButtons">
      <a href="javascript:document.frm.submit();" class="bo_Button80">Submit</a>
      <a href="javascript:window.close();" class="bo_Button80">Cancel</a>
    </span>
    
  </form>
  
    
</body>
</html>
