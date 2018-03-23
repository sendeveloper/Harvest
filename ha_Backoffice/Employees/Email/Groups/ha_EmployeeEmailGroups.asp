<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include virtual="includes/connection.asp"-->
<!--#include virtual="ha_BackOffice/includes/Config.asp"-->

<%
	ColorTab = 4
	PageHeading = "Employee E-mail Groups"
%>

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<title>Harvest American Backoffice - <%=PageHeading%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    
	<link href="<%=strPathIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet" />
	<link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>menu/dropdowntabfiles/halfmoontabs.css" />
	<script type="text/javascript" src="<%=strPathIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>
  
    <style type="text/css">
		td
			{
			vertical-align:	top;
			}
			
		th
			{
			text-align:		left;
			border-bottom:	1px solid black;
			}
			
		div.members
			{
			border:			1px solid gray;
			height:			5em;
			}
    </style>
	
</head>

<body class="gray_desktop">
  <table cellpadding="0" cellspacing="0" class="MainBody" style="margin: 20px auto 0;width: 1200px !important;">
    <tr>
      <td>
        <!--#include virtual="ha_Backoffice/includes/bodyparts/ha_PageHeading.inc"-->
      </td>
    </tr>
	
	<tr><td>&nbsp;</td></tr>
	
    <tr>
      <td align="center">
        <table width="95%" cellpadding="2" cellspacing="2">
		  <tr>
		    <td colspan="4">
			  We use E-mail groups as a shortcut for reaching our staff internally.
			  Each of our E-mail groups are set to forward to specific addresses as shown below.
			  <br><br>
			  CPanel is our E-mail platform that we rent from QTH.com.
			  On CPanel they are called "Forwarders" which is where you will find the editing capabilities.
			</td>
		  </tr>
		  
		  <tr><td>&nbsp;</td></tr>
		  <tr><td>&nbsp;</td></tr>
		  
		  <tr>
			<th width="15%">Group Name</th>
			<th width="25%">E-mail Address</th>
			<th width="40%">Description</th>
			<th width="20%" style="Text-Align: Center;">Action</th>
		  </tr>
		  
	<%
		set rs=server.createObject("ADODB.Recordset")
		SQL = "ha_Employee_EmailGroups_list"
		rs.Open SQL, connLocal, 3, 3, 4
		
		
		set rsMembersList=server.createObject("ADODB.Recordset")

		if not RS.EOF then
			do while not rs.eof
	%>
		
		  <tr>
			<td style="<%=detailStyle%>">
			  <b><%=rs("GroupName")%></b>
			</td>
			<td style="<%=detailStyle%>">
			  <%=rs("EmailAddress")%>
			</td>
			<td style="<%=detailStyle%>">
			  <%=rs("GroupDescription")%>
			</td>
			<td style="<%=detailStyle%>" align="center">&nbsp;
			  
			</td>
		  </tr>
		  
		  <tr>
			<td style="<%=detailStyle%>" align="right">
			  Members:
			</td>

			<td style="<%=detailStyle%>" colspan="2">
			  <div class="members">
              <%
					SQL = "ha_Employee_EmailGroup_Members_list(" & rs("ID") & ")"
					rsMembersList.Open SQL, connLocal, 3, 3, 4
					if not rsMembersList.EOF then
					response.Write(rsMembersList("Members_list"))					
					end if
					rsMembersList.close
					
			  %>
			    &nbsp;
			  </div>
			</td>

			<td style="<%=detailStyle%>" align="center">
			  <a href="javascript:void(0);" onclick="window.open('ha_EmployeeEmailGroups_edit.asp?tMode=Edit&ID=<%=rs("ID")%>&GroupName=<%=rs("GroupName")%>','_blank','width=10,height=10')">
	              Edit
               </a>
			</td>
		  </tr>
		  
		  <tr><td>&nbsp;</td></tr>
		
	<%
				rs.MoveNext
			loop
		end if
		
		rs.close
		connLocal.close
	%>
		  
		
		  <tr>
		    <td colspan="4">
			  <hr>
			</td>
		  </tr>
		  
		</table>
      </td>
    </tr>
	
	<tr><td>&nbsp;</td></tr>
	<tr><td>&nbsp;</td></tr>

  </table>
  <!--#include virtual="ha_BackOffice/includes/BodyParts/ha_PageTail.inc"-->
</body>
</html>

