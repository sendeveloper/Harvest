<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/Config.asp"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsHelp.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->
  <!--#include virtual="ha_backoffice/includes/Dates/ConvDate.inc"-->
  
<%

	If Request("a") = "Attach" Then
		strSQL = "ha_Employee_Permission_attach(" & Request("ID") & ", " & Request("fid") & ", '" & Session("ha_shortname") & "')"
		'response.write strSQL & "<br>"
		connCasper10.Execute strSQL
	End If
	
	If Request("a") = "Remove" Then
		strSQL = "ha_Employee_Permission_Remove(" & Request("ID") & ", " & Request("fid") & ", '" & Session("ha_shortname") & "')"
		connCasper10.Execute strSQL
	End If
	
	If Request("a") = "Save" Then
		strSQL = "ha_Employee_Permission_Save(" & Request("ID") & ", '" & Request("p") & "', '" & Session("ha_shortname") & "')"
		connCasper10.Execute strSQL
		response.redirect "ha_EmployeePermissions_edit.asp?ID=" & Request("ID")
	End If
	
	Dim rs: Set rs = Server.CreateObject("ADODB.Recordset")

		ID = Request("ID")
		Title = "Employee Permissions"
		
		strSQL = "SELECT * FROM ha_EmployeesFormatted WHERE ID = " & ID
        rs.open strSQL, connCasper10, 0, 1
		
		eFullName			= Null2Space(rs("FullName"))
		eUserName			= Null2Space(rs("UserName"))
		ePassword			= Null2Space(rs("Password"))

		eCreatedBy			= Null2Space(rs("CreatedBy"))
		eCreatedDate		= Null2Space(rs("CreatedDate"))
		eEditedBy			= Null2Space(rs("EditedBy"))
		eEditedDate			= Null2Space(rs("EditedDate"))
		
        rs.close
		
%>

<html>
<head>
	<title><%=Title%></title>

	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
  
	<link href="/ha_Backoffice/includes/ha_BackOfficePopup.css" type="text/css" rel="stylesheet">
	<script language="javascript" type="text/javascript" src="/includes/haFunctions.js"></script>
	
	<script type="text/javascript">
	    var httpFiles = new XMLHttpRequest();
		var rowCount = 0;
		var permissionText  = new Array();
		var paramsStart = '';
		var paramsNow = '';
		
		var buttonsClose = '<span class="popupButtons">' +
			'	<a href="javascript:clickEditClose();" class="bo_Button80">Close</a>' +
			'</span>'
		var buttonsCancelSave = '<span id="spanCancelSave" class="popupButtons">' +
			'	<a href="javascript:clickEditClose();" class="bo_Button80">Cancel</a>' +
			'	<a href="javascript:clickEditSave();" class="bo_Button80">Save</a>' +
			'</span>'
		
		
		function changeEditDropDown()
			{
			paramsNow = checkEditValues()
			//alert(paramsNow);
			
			if (paramsNow == paramsStart)
				{
				document.getElementById('editButtons').innerHTML = buttonsClose;			
				}
			else
				{
				document.getElementById('editButtons').innerHTML = buttonsCancelSave;			
				}
			}
			
		function changeQualify(b)
			{
			}
		
		function checkEditValues()
			{
			var params = '';

			//Read the values of all the select boxes
			for(i=0; i<document.formEdit.elements.length; i++)
				{	
				if (document.formEdit.elements[i].type == 'select-one')
					{
					var fieldName = document.formEdit.elements[i].name;
					//alert(i + ' | ' + document.formEdit.elements[i].name + ' | ' + document.formEdit.elements[i].value);
					var fieldValue = document.formEdit.elements[i].value;

					params += fieldName + '~' + fieldValue + '|';
					}
				}
				return params;
			}

		function clickAdd()
			{
			document.getElementById('divAdd').style.visibility = "visible";
			}	
			
		function clickAddClose()
			{
			document.getElementById('divAdd').style.visibility = "hidden";
			}	
			
		function clickAttachSelect(fid)
			{
			clickAddClose();
			var URL = 'ha_EmployeePermissions_edit.asp?a=Attach&fid=' + fid + '&ID=<%=ID%>';
			//alert(URL);
			window.location.href = URL;
			}	
			
		function clickEdit(id, desc, body)
			{
			document.getElementById('editSubhead').innerHTML = desc;
			document.getElementById('editBody').innerHTML = body.replace(/{/g,'"');
			document.getElementById('editButtons').innerHTML = buttonsClose;			
			document.getElementById('divEdit').style.visibility = "visible";
			
			paramsStart = checkEditValues()
			//alert(paramsStart);
			}	
			
		function clickEditClose()
			{
			document.getElementById('divEdit').style.visibility = "hidden";
			}	

			
		function clickEditSave()
			{
			document.getElementById('divEdit').style.visibility = "hidden";
			
			var URL = 'ha_EmployeePermissions_edit.asp?a=Save&ID=<%=ID%>&p=' + paramsNow;
			//alert(URL);
			window.location.href = URL;
			}	
			
  		function clickRemove(fid)
			{
			var URL = 'ha_EmployeePermissions_edit.asp?a=Remove&fid=' + fid + '&ID=<%=ID%>';
			window.location.href = URL;
			}	
			
		function clickPermissionDefinitionsSearch()
			{
			var f = document.getElementById('pdSearchText').value;
			if (f > '')
				{
				getPermissionDefinitions(f);
				}
			}	
			
		function getPermissionDefinitions(f) 
			{		
			var url = '/ha_BackOffice/includes/ajax/ha_PermissionDefinitions_search.asp' +
				'?f=' + escape(f) +
				'&Now=' + escape(Date());
			
			httpFiles.open('GET', url, true);
			httpFiles.onreadystatechange = getPermissionDefinitionsResponse;
			httpFiles.send();
			}

		function getPermissionDefinitionsResponse() 
			{
			if (httpFiles.readyState == 4) 
				{
				if (httpFiles.status == 200) 
					{
					document.getElementById('divPermissionDefinitions').innerHTML = httpFiles.responseText;
					}
				}
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
			
</head>

<body onLoad="SetScreen(1100,650);">
  <form method="Post" action="ha_EmployeePermissions_post.asp?ID=<%=ID%>" name="frm">

	  <span class="popupHeading"><%=Title%></span>
	  
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">	  
  <tr>
    <td colspan="2">
	
      <table width="95%" border="0" cellspacing="2" cellpadding="1" align="center">

		<tr><td width="35%>&nbsp;</td><td width="65%>&nbsp;</td></tr>
			
		<tr>
		  <td align="right" style="padding-bottom: 8px;">
			Name:
		  </td>
		  <td style="padding-bottom: 8px;">
		    <b><%=eFullName%></b>
		  </td>		 
		</tr>
		
		<tr>
		  <td align="right">
			Username:
		  </td>
		  <td>
			<input type="text" name="StaffUN" id="StaffUN" 
				size="20" value="<%=iif(HasPermission("Read", 7), eUserName, Bullets8)%>"
				<%=iif(HasPermission("Edit", 7),""," Disabled")%>>
		  </td>		 
		</tr>

		<tr>
		  <td align="right">
			Password:
		  </td>
		  <td>
			<input type="text" name="StaffPWD" id="StaffPWD"
				size="20" value="<%=iif(HasPermission("Read", 8), ePassword, Bullets8)%>"
				<%=iif(HasPermission("Edit", 8),""," Disabled")%>>
		  </td>		 
		</tr>

		<tr><td>&nbsp;</td></tr>
	  </table>
	  
      <table width="95%" border="0" cellspacing="2" cellpadding="1" align="center">
		
		<tr>
		  <td align="right" colspan="7">
   		    <a href="javascript:clickAdd();">Add Permission</a>
		  </td>
		</tr>
		
		<tr><td>&nbsp;</td></tr>
		
		<tr>
		  <td width="5%" class="head">Row</td>
		  <td width="10%" class="head">Program</td>
		  <td width="10%" class="head">Category</td>
		  <td width="20%" class="head">Page</td>
		  <td width="35%" class="head">Description</td>
		  <td width="10%" class="head">Permissions</td>		  
		  <td width="10%" class="head">Actions</td>
		</tr>

<%
        Dim nRecordsPerPage, icurrentPage, intPageCount, intRecordCount, iRecordperpage        

		strSQL = "ha_Employee_Permissions_list(" & cStr(ID) & ", " & Session("ha_EmployeeID") & ")"
		'response.write strSQL

		rs.CursorLocation = 3
		rs.Open strSQL, connCasper10, 3, 3, 4
		iRecordperpage = 20
		
		'------------------------------------Starting Paging from Here-----------------------------------
		SeparatorRows = 3
		SeparatorWidth = "1px"
		SeparatorColor = "#C0C0C0"
%>
		<!--#include virtual="ha_backoffice/includes/FunctionsPageStart.inc"--> 
		
		<tr id="row[<%=row%>]">
				  
		  <td class="row" style="<%=detailStyle%>" align="center">
			<%=row%>
		  </td>
		  <td class="row" style="<%=detailStyle%>">
		    <%=rs("Program")%>
		  </td>
		  <td class="row" style="<%=detailStyle%>">
		    <%=rs("Category")%>
		  </td>
		  <td class="row" style="<%=detailStyle%>">
		    <%=rs("Page")%>
		  </td>
		  <td class="row" style="<%=detailStyle%>">
		    <%=rs("Description")%>
		  </td>
		  <td class="row" style="<%=detailStyle%> font-family: Courier;" align="center"
            onMouseOver='altOnLeft(this, "Permissions", "<%=rs("Description")%>", "<%=rs("PermissionText")%>", event);'
            onMouseOut='altOff();'>
		    <%=rs("PermissionString")%>
		  </td>
		  <td class="row" style="<%=detailStyle%>" align="center">
   		    <a href="javascript:clickEdit(<%=rs("ID")%>, '<%=rs("Description")%>', '<%=rs("PermissionEditText")%>');">Edit</a>
   		    <a href="javascript:clickRemove(<%=rs("PermissionID")%>);">Remove</a>
		  </td>
		</tr>
		
		<!--#include virtual="ha_backoffice/includes/FunctionsPageEnd.inc"-->
		
	  </table>
    </td>
  </tr>
</table>

    <span class="popupCredit"><%=CreditLine(eCreatedBy, eCreatedDate, eEditedBy, eEditedDate)%></span>	

    <span class="popupButtons">
      <a href="javascript:document.frm.submit();" class="bo_Button80">Submit</a>
      <a href="javascript:window.close();" class="bo_Button80">Close</a>
    </span>

  </form>
  <!--#include virtual="ha_backoffice/includes/FunctionsMouseover.inc"-->
  
<!-- In-page Popup Form -->
<div id="divAdd" name="divAdd" class="divEdit">
  <form action="javascript:whichKey();">
	<table width="100%" cellspacing="5" cellpadding="5">
	  <tr>
	    <td width="100%">
		  <span style="margin: auto; 
			font-weight: bold; 
			font-size: 16pt; 
			display: block;
			text-align: center;
			border-bottom: 1px solid black;">
			Add Permission
		  </span>
		</td>
	  </tr>
	  
	  <tr>
		<td>
		  <input type="text" id="pdSearchText" name="pdSearchText" size="35" value="">
		  <a href="javascript:clickPermissionDefinitionsSearch();" class="bo_Button80">Search</a>
		</td>
	  <tr>
	  
	  <tr>
	    <td width="100%">
		  <div id="divPermissionDefinitions" name="divPermissionDefinitions"
			style="border: 1px solid black; 
		    height: 300px; 
		    background-color: white;
			overflow-y: scroll;">
		  
		  </div>
		</td>
	  </tr>
	    
	</table>
	
    <span class="popupButtons">
      <a href="javascript:clickAddClose();" class="bo_Button80">Close</a>
    </span>
  </form>
</div>

<div id="divEdit" name="divEdit" class="divEdit">
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
			Permission Edit
		  </span>
		</td>
	  </tr>
	  
	  <tr>
		<td style="width: 30%; text-align: right;">
		  Name:
		</td>
		<td style="width: 70%; font-weight: bold;">
		  <%=eFullName%>
		</td>
	  <tr>
	  
	  <tr>
		<td style="text-align: right;">
		  Description:
		</td>
		<td style="font-weight: bold;">
		  <span id="editSubhead">Subhead Goes Here</span>
		</td>
	  </tr>
	    
	  <tr>
	    <td colspan="2" style="padding: 10px 0 20px 0;">
		  <hr style="width: 90%; text-align: center; border: none; border: none; height: 1px;">
		  <span id="editBody">Body Goes Here</span>
		</td>
	  </tr>
	  
	</table>
	
	<span id="editButtons">&nbsp;</span>
	
  </form>
</div>
  
</body>
</html>
