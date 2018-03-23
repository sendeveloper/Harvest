<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

  <!--#include virtual="includes/connection.asp"-->  
  
<%
	Title = "&quot;To Do&quot; Rules"
%>

<html>
<head>
  <title><%=Title%></title>

  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
  
  <link href="/ha_Backoffice/includes/ha_BackOfficePopup.css" type="text/css" rel="stylesheet">
  <script language="javascript" type="text/javascript" src="/includes/haFunctions.js"></script>
  
  <style type="text/css">
	
	td.col1
		{
		vertical-align:	top;
		width:			20%;
		text-align:		right;
		padding-right:	1em;
		font-weight:	bold;
		}
  
	td.col2
		{
		vertical-align:	top;
		width:			80%;
		align:			left;
		}
		
	th
		{
		font-size:		11pt;
		text-align:		left;
		border-bottom: 	1px solid black;
		}
		
  </style>

  <script type="text/javascript">
  			
	function formLoad()
		{
		SetScreen(1000,950);
		}
		
  </script>
</head>

<body onLoad="formLoad();">

	  <span class="popupHeading"><%=Title%></span>
	  
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">	  
  <tr>
    <td>
	
      <table width="95%" border="0" cellspacing="2" cellpadding="1" align="center">
		
		<tr><td>&nbsp;</td></tr>
		
		<tr>
		  <td colspan="2">
		    Here are rules everybody should know about "To Do"s.
		  </td>
		</tr>
		
		<tr><td>&nbsp;</td></tr>

	  </table>
		
      <table width="95%" border="0" cellspacing="2" cellpadding="1" align="center">
	  	  
		<!-- .................... Adding .................... -->
  		<tr>
		  <th colspan="2">
		    Adding New "To Do"s
		  </th>
		</tr>
		
		<tr>
		  <td class="col1">
		    Priority
		  </td>
		  <td class="col2">
			Leave blank
		  </td>
		</tr>

		<tr>
		  <td class="col1">
		    Title
		  </td>
		  <td class="col2">
			Assign the task a title.
		  </td>
		</tr>

		<tr>
		  <td class="col1">
		    Description
		  </td>
		  <td class="col2">
			Brief summary of the work needed. 
		  </td>
		</tr>
		
		<tr>
		  <td class="col1">
		    Server Name
		  </td>
		  <td class="col2">
			Fill in if applicable.
		  </td>
		</tr>

		<tr>
		  <td class="col1">
		    Category
		  </td>
		  <td class="col2">
			Fill in if known.
		  </td>
		</tr>

		<tr>
		  <td class="col1">
		    Status
		  </td>
		  <td class="col2">
			Leave blank.
		  </td>
		</tr>

		<tr>
		  <td class="col1">
		    Notes
		  </td>
		  <td class="col2">
			Outline work needed, add specifics, notes, email conversations, etc. 
			Update this field with date and name of person updating periodically.
		  </td>
		</tr>

		<tr>
		  <td class="col1">
		    Assigned & Managed By
		  </td>
		  <td class="col2">
			Leave blank or fill in if known with certainty (assigning to self).
		  </td>
		</tr>
		
		<tr><td>&nbsp;</td></tr>
		
		<!-- .................... Editing .................... -->
  		<tr>
		  <th colspan="2">
		    Editing Existing "To Do"s
		  </th>
		</tr>

		<tr>
		  <td class="col1">
		    Priority
		  </td>
		  <td class="col2">
			Do not change unless during "To Do" Meeting.
		  </td>
		</tr>
		  
		<tr>
		  <td class="col1">
		    Title
		  </td>
		  <td class="col2">
			Make clarifications only if needed.
		  </td>
		</tr>
		
		<tr>
		  <td class="col1">
		    Description
		  </td>
		  <td class="col2">
			Make clarifications only if needed.
		  </td>
		</tr>
		
		<tr>
		  <td class="col1">
		    Server Name
		  </td>
		  <td class="col2">
			Can update if needed.
		  </td>
		</tr>
		
		<tr>
		  <td class="col1">
		    Category
		  </td>
		  <td class="col2">
			Can update if needed.
		  </td>
		</tr>
		
		<tr>
		  <td class="col1">
		    Status
		  </td>
		  <td class="col2">
			Update as work progresses. Mark "Final Check" if you believe the task is done. 
			Do NOT mark as complete unless during "To Do" meeting.
		  </td>
		</tr>

		<tr>
		  <td class="col1">
		    Notes
		  </td>
		  <td class="col2">
			Update this field with date and name of person doing updating periodically. 
			Include correspondence and other valuable information.
		  </td>
		</tr>
		
		<tr>
		  <td class="col1">
		    Assigned & Managed By
		  </td>
		  <td class="col2">
			Do not change unless during "To Do" meeting.
		  </td>
		</tr>
		
		<tr><td>&nbsp;</td></tr>
		
		<!-- .................... Agenda .................... -->
  		<tr>
		  <th colspan="2">
		    "To Do" Meeting Agenda
		  </th>
		</tr>

		<tr>
		  <td class="col1">
		    1.)
		  </td>
		  <td class="col2">
			Review "To Do"s with a status of Blank and New.  Assign staff and set priorities.
		  </td>
		</tr>
		
		<tr>
		  <td class="col1">
		    2.)
		  </td>
		  <td class="col2">
			Set "To Do Page Filters: Priority of Blank, 1, 2, and 3; Status of everything except Done and On Hold.
			Then we go around the table and let each individual discuss their projects.
		  </td>
		</tr>
		
		<tr>
		  <td class="col1">
		    3.)
		  </td>
		  <td class="col2">
			Review All "To Do"s with a Priority of 1.
		  </td>
		</tr>

		<tr>
		  <td class="col1">
		    4.)
		  </td>
		  <td class="col2">
			Check Statistics and note the total number of "To Do"s.
		  </td>
		</tr>

		<tr>
		  <td class="col1">
		    5.)
		  </td>
		  <td class="col2">
			Acknowledge accomplishments.
		  </td>
		</tr>
		
		<tr><td>&nbsp;</td></tr>

	  </table>
					
      <table width="95%" border="0" cellspacing="2" cellpadding="1" align="center">
					
		<tr>
		  <td style="border-bottom: 2px solid gray;">
			&nbsp;
		  </td>
		</tr>
		
		<tr><td>&nbsp;</td></tr>
		
	  </table>
    </td>
  </tr>
</table>

    <span class="popupButtons" style="padding-bottom: 3em;">
      <a href="javascript:window.close();" class="bo_Button80">Close</a>
    </span>

</body>
</html>
