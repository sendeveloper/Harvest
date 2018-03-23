<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/Config.asp"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsHelp.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->
  <!--#include virtual="ha_backoffice/includes/Dates/ConvDate.inc"-->
  
<%
    Dim rs, strSQL, ID
	Set rs = Server.CreateObject("ADODB.Recordset")
   

	ID 				= 0
	eGroupName		= Request("GroupName")
	eCreatedBy		= ""
	eCreatedDate	= ""
	eEditedBy		= ""
	eEditedDate		= ""

    If Request("ID") = "" Then
		Title = "Add To Employee Email Groups"
	Else
		Title = "Edit To Employee Email Groups"
		
		ID = Request("ID")
		

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
    var httpFiles = new XMLHttpRequest();

		
	function formLoad()
		{
		SetScreen(1100,780);
		}			
    </script>
  

    
    <style type="text/css">
			
    tr
		{
		vertical-align: top;
		text-align: left;
	   
		}
	.StaffEmailTD
	{
		width: 50%;		
	}
		
	.StaffInputTD
		{
		width: 50%;
	    padding-left: 28%;
			
		}	
			
	.StaffEmail
		{
		
		text-align:left;
	    line-height: 20.9px;
		}
	
    </style>
  
</head>

<body onLoad="formLoad();">
  <form method="Post" action="ha_EmployeeEmailGroups_post.asp" name="frm">
    <input type="hidden" id="GroupID" name="GroupID" value="<%=ID%>">
    

	  <span class="popupHeading"><%=Title%></span>
	  
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">	  
  <tr>
    <td colspan="2">
	
      <table width="100%" border="0" cellspacing="2" cellpadding="2" align="center">

		
			
        
         
		
		<tr vAlign="top">
        
          <td>
          	
            <table width="100%" border="0" >
            
             	<tr vAlign="top">
					<td  align="center" style="padding-left: 35%;">
			  			<b><%=eGroupName%></b>
					</td>
		 	    </tr>

			  <tr>
			    	<td vAlign="top" style="width:50%; padding-left:28%;">
				 	 <span style="font-weight: bold; text-decoration: underline; ">Member</span>
                  	</td>

                    <td vAlign="top" style="width:50%">
					  <span style="font-weight: bold;  text-decoration: underline;">Email</span>
                    </td>

             </tr>
                  
<%
					strSQL = "ha_Employee_EmailGroups_Members_list(" & Request("ID") & ")"
					rs.Open strSQL, connLocal, 3, 3, 4
					If not rs.eof Then
						While not rs.eof
							Response.write "<tr class=""MembersTd"">"  & rs("HTMLGroupMember") &_
							  rs("HTMLMemberEmail") & "</tr>"
							rs.MoveNext
						Wend
					End If
					rs.Close
%>				  
				<!--/td>
			    <td vAlign="top" style="width:50%">
				  <br>
				  <span style="font-weight: bold;  text-decoration: underline;">Email</span><br><br>
<%
					'strSQL = "ha_Employee_EmailGroups_Members_list(" & Request("ID") & ")"
					'rs.Open strSQL, connLocal, 3, 3, 4
					'If not rs.eof Then
					'	While not rs.eof
					'		Response.write rs("HTMLMemberEmail") & "<br>" & vbCrLf
					'		rs.MoveNext
					'	Wend
					'End If
					'rs.Close
%>				  
				</td>
			  </tr-->
			</table>
          
          </td>
		  
		</tr>
		
		
		
		
		
      </table>
    </td>
  </tr>
</table>

<%
		strSQL = "ha_Employee_EmailGroup_Members_credits(" & Request("ID") & ")"
		rs.Open strSQL, connLocal, 3, 3, 4
					If not rs.eof Then
						eCreatedBy 		= rs("CreatedBy")
						eCreatedDate 	= rs("CreatedDate")
						eEditedBy 		= rs("EditedBy")
						eEditedDate 	= rs("EditedDate")
		
%>

    <span class="popupCredit"><%=CreditLine(eCreatedBy, eCreatedDate, eEditedBy, eEditedDate)%></span>
 <%
 End If
 %>

    <span class="popupButtons">
      <a href="javascript:document.frm.submit();" class="bo_Button80">Submit</a>
      <a href="javascript:window.close();" class="bo_Button80">Close</a>
    </span>

    	<div id="errorcon">&nbsp;</div>

  </form>

  
</body>
</html>
