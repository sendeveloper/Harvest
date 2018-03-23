<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsHelp.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->
  
<%
	Dim rs: Set rs = Server.CreateObject("ADODB.Recordset")

	If IsNull(Request("ID")) or Request("ID")=0 Then
		ID = 0
		Title = "Job Description Titles Add"

		eDate = ""
		eDescription = ""
	Else
		ID = Request("ID")
		Title = "Job Description Titles Edit"
		
		strSQL = "SELECT * FROM ha_JobDescriptionTitles WHERE ID = " & ID
        rs.open strSQL, connLocal, 0, 1

		eCategory 		= Null2Space(rs("Category"))
		eTitleName 		= Null2Space(rs("TitleName"))
		ePurpose 		= Null2Space(rs("TitlePurpose"))
		eCreatedBy		= Null2Space(rs("CreatedBy"))
		eCreatedDate	= Null2Space(rs("CreatedDate"))
		eEditedBy		= Null2Space(rs("EditedBy"))
		eEditedDate		= Null2Space(rs("EditedDate"))
		
        rs.close
	End If
	
%>

<html>
<head>
  <title>Harvest American Backoffice - <%=Title%></title>
  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
  
  <link href="/ha_Backoffice/includes/ha_BackOfficePopup.css" type="text/css" rel="stylesheet">
  <script language="javascript" type="text/javascript" src="/includes/haFunctions.js"></script>
  <style type="text/css">
			
    input[type="text"]
		{
		width: 55em;   
		}
	
    select
		{
		width: 12.5em;   
		}
	
    textarea 
		{
		font-family:    Verdana, Arial, Helvetica, sans-serif;
		width: 			55em;   
		height: 		5em; 
		resize: 		none; 
		}
		
    .divDisplay
		{
		width: 			55em;   
		border:			1px solid #A9A9A9;
		overflow: 		auto;
		padding: 		2px;
		}

	.divAdd
		{
		position:absolute; 
		border: 2px solid black;
        top: 90px; 
       	left: 180px; 
        width: 750px; 
       	height: 450px;
        background-color: #FFFFCC;
       	visibility: hidden;
		}
		
  </style>
  
  <script type="text/javascript">
    var httpFiles = new XMLHttpRequest();
  				
	function clickAddClose()
		{
		document.getElementById('divAdd').style.visibility = "hidden";
		}	
		
	function clickAddSearch()
		{
		var f = document.getElementById('pdSearchText').value;
		if (f > '')
			{
			if (document.getElementById('divAddTitle').innerHTML == 'Add Assignee')
				{
				getAssigneeList(f);
				}
			else
				{
				getTaskList(f);
				}
			}
		else
			{
			document.getElementById('divAddList').innerHTML = '';
			}
		}	
		
  	function clickAssigneeAdd()
		{
		document.getElementById('divAddTitle').innerHTML = 'Add Assignee';
		document.getElementById('divAddList').innerHTML = '';
		getAssigneeList('');
		document.getElementById('divAdd').style.visibility = "visible";
		}
		
  	function clickAsssigneeRemove(id)
		{
		document.frm.SaveAction.value = 'AssigneeRemove';
		document.frm.PassID.value = id;
		document.frm.submit();
		}
		
	function clickAssigneeSelect(fid)
		{
		clickAddClose();
		document.frm.SaveAction.value = 'AssigneeAttach';
		document.frm.PassID.value = fid;
		document.frm.submit();
		}		
					
	function getAssigneeList(f) 
		{		
		var url = '/ha_BackOffice/includes/ajax/ha_Employee_search.asp' +
			'?f=' + escape(f) +
			'&Now=' + escape(Date());
		
		httpFiles.open('GET', url, true);
		httpFiles.onreadystatechange = getAssigneeResponse;
		httpFiles.send();
		}
		
	function getAssigneeResponse() 
		{
		if (httpFiles.readyState == 4) 
			{
			if (httpFiles.status == 200) 
				{
				document.getElementById('divAddList').innerHTML = httpFiles.responseText;
				}
			}
		}
				
  	function clickTaskAdd()
		{
		document.getElementById('divAddTitle').innerHTML = 'Add Task';
		document.getElementById('divAddList').innerHTML = '';

		document.getElementById('divAdd').style.visibility = "visible";
		}
		
  	function clickTaskMoveDown(id)
		{
		document.frm.SaveAction.value = 'TaskMoveDown';
		document.frm.PassID.value = id;
		document.frm.submit();
		}
		
  	function clickTaskMoveUp(id)
		{
		document.frm.SaveAction.value = 'TaskMoveUp';
		document.frm.PassID.value = id;
		document.frm.submit();
		}

	function clickTaskRemove(id)
		{
		document.frm.SaveAction.value = 'TaskRemove';
		document.frm.PassID.value = id;
		document.frm.submit();
		}

	function clickTaskSelect(fid)
		{
		clickAddClose();
		document.frm.SaveAction.value = 'TaskAttach';
		document.frm.PassID.value = fid;
		document.frm.submit();
		}		
		
	function clickTaskSearch()
		{
		var f = document.getElementById('pdSearchText').value;
		if (f > '')
			{
			getTaskList(f);
			}
		}	
			
	function getTaskList(f) 
		{		
		var url = '/ha_BackOffice/includes/ajax/ha_Task_search.asp' +
			'?f=' + escape(f) +
			'&Now=' + escape(Date());
		
		httpFiles.open('GET', url, true);
		httpFiles.onreadystatechange = getTaskResponse;
		httpFiles.send();
		}
		
	function getTaskResponse() 
		{
		if (httpFiles.readyState == 4) 
			{
			if (httpFiles.status == 200) 
				{
				document.getElementById('divAddList').innerHTML = httpFiles.responseText;
				}
			}
		}
		
	function formLoad()
		{
		SetScreen(1100,700);
		}
				
  </script>
  
</head>

<body onLoad="formLoad();">
  <form method="Post" action="ha_JobDescriptionTitles_post.asp" name="frm">
    <input type="hidden" id="TitleID" name="TitleID" value="<%=ID%>">
    <input type="hidden" id="PassID" name="PassID" value="<%=ID%>">
    <input type="hidden" id="SaveAction" name="SaveAction" value="SaveFields">

	  <span class="popupHeading"><%=Title%></span>
	  
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">	  
  <tr>
    <td colspan="2">
	
      <table width="95%" border="0" cellspacing="2" cellpadding="2" align="center">

		<tr><td>&nbsp;</td></tr>
			
		<tr>
		  <td width="15%" align="right">
			Category:
		  </td>
		  <td colspan="2">
			<Select id="Category" name="Category">
  
<%
					strSQL = " util_HTML_option_list('ha_types_repl', 'class', 'JobDescriptions - TitleCategories', 'value', 'value', 'sequence', '', '" & eCategory & "')"
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
		  </td>
		</tr>

		<tr>
		  <td align="right">
			Title Name:
		  </td>
		  <td colspan="2">
			<input type="text" name="TitleName" id="TitleName"
				size="65" value="<%=eTitleName%>">				
		  </td>		 
		</tr>

		<tr vAlign="top">
		  <td align="right">
			Purpose:
		  </td>
		  <td colspan="2">
			<textarea id="Purpose" name="Purpose"><%=ePurpose%></textarea>
		  </td>
		</tr>

		<tr vAlign="top">
		  <td align="right">
			Assigned To:
		  </td>
		  <td style="display: inline-block;">
			<div id="divAssignees" name="divAssignees" class="divDisplay" style="height: 5em;">
			  <table border="0" cellspacing="2" cellpadding="0">
			
<%
					strSQL = "[ha_JobDescription_Title/Employee_list](" & ID & ")"
					rs.Open strSQL, connLocal, 3, 3, 4
					If not rs.eof Then
						While not rs.eof
							s = "<tr>"
							s = s & "<td width='100'>" & rs("ShortName") & "</td>"
							s = s & "<td width='340'>" & rs("FullName") & "</td>"
							s = s & "<td width='100'><a href='javascript:clickAsssigneeRemove(" & rs("ID") & ");' class='bo_Button60'>Remove</a>"
							s = s & "</tr>"
							Response.write  s
							rs.MoveNext
						Wend
					End If
					rs.Close
%>
			  </table>
			</div>
		  </td>
		  <td style="display: inline-block;">
			<a href="javascript:clickAssigneeAdd();" class="bo_Button60">Add</a>
		  </td>
		</tr>
		
		<tr vAlign="top">
		  <td align="right">
			Tasks:
		  </td>
		  <td style="display: inline-block;">
			<div id="divTasks" name="divTasks"  class="divDisplay" style="height: 15em;">
			  <table border="0" cellspacing="2" cellpadding="0">
			
<%
					strSQL = "[ha_JobDescription_Title/Task_list](" & ID & ")"
					rs.Open strSQL, connLocal, 3, 3, 4
					If not rs.eof Then
						While not rs.eof
							s = "<tr>"
							s = s & "<td width='100'>" & rs("Category") & "</td>"
							s = s & "<td width='340'>" & rs("TaskName") & "</td>"
							s = s & "<td width='210'>"
							s = s & "  <a href='javascript:clickTaskRemove(" & rs("ID") & ");' class='bo_Button60'>Remove</a>"
							s = s & "  <a href='javascript:clickTaskMoveUp(" & rs("ID") & ");' class='bo_Button60'>Move Up</a>"
							s = s & "  <a href='javascript:clickTaskMoveDown(" & rs("ID") & ");' class='bo_Button60'>Move Dn</a>"
							s = s & "</td></tr>"
							Response.write  s
							rs.MoveNext
						Wend
					End If
					rs.Close
%>
			  </table>
			</div>
		  </td>
		  <td style="display: inline-block;">
			<a href="javascript:clickTaskAdd();" class="bo_Button60">Add</a>
		  </td>
		</tr>
		
		<tr><td>&nbsp;</td></tr>
		
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

<!-- In-page Popup Form -->
<div id="divAdd" name="divAdd" class="divAdd">
  <form action="javascript:whichKey();">
	<table width="100%" cellspacing="5" cellpadding="5">
	  <tr>
	    <td width="100%">
		  <span id="divAddTitle" style="margin: auto; 
				font-weight: bold; 
				font-size: 16pt; 
				display: block;
				text-align: center;
				border-bottom: 1px solid black;">
			&nbsp;
		  </span>
		</td>
	  </tr>
	  
	  <tr>
		<td>
		  <input type="text" id="pdSearchText" name="pdSearchText" style="width: 30em;" value="">
		  <a href="javascript:clickAddSearch();" class="bo_Button80">Search</a>
		</td>
	  <tr>
	  
	  <tr>
	    <td width="100%">
		  <div id="divAddList" name="divAddList"
			style="border: 1px solid black; 
				width: 50em;
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
  
</body>
</html>

