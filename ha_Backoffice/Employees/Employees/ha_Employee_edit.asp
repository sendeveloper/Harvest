<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsDates.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsHelp.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsPermissions.inc"-->
  <!--#include virtual="ha_backoffice/includes/Dates/ConvDate.inc"-->
  
<%
	Dim rs: Set rs = Server.CreateObject("ADODB.Recordset")

	If IsNull(Request("ID")) or Request("ID")=0 Then
		ID = 0
		Title = "Employee Add"

		eFirstName			= ""
		eLastName			= ""
		eEmail				= ""
		eDOB				= ""
		eBirthday			= ""
		eAddress1			= ""
		eAddress2			= ""
		eCity				= ""
		eState				= ""
		eZipCode			= ""
		eWorkLocation		= ""
		eHomePhone			= ""
		eMobilePhone		= ""
		eSkype				= ""
		eHarvestID			= ""
		
		eCreatedBy			= ""
		eCreatedDate		= ""
		eEditedBy			= ""
		eEditedDate			= ""

	Else
		ID = Request("ID")
		Title = "Employee Edit"
		
		strSQL = "SELECT * FROM ha_EmployeesFormatted WHERE ID = " & ID
        rs.open strSQL, connCasper10, 0, 1
		
		eFirstName			= Null2Space(rs("FirstName"))
		eLastName			= Null2Space(rs("LastName"))
		eEmail				= Null2Space(rs("Email"))
		eDOB				= Null2Space(rs("DOB"))
		eBirthday			= Null2Space(rs("Birthday"))
		eAddressLine1		= Null2Space(rs("AddressLine1"))
		eAddressLine2		= Null2Space(rs("AddressLine2"))
		eCity				= Null2Space(rs("City"))
		eState				= Null2Space(rs("State"))
		ePostalCode			= Null2Space(rs("PostalCode"))
		eWorkLocation		= Null2Space(rs("WorkLocation"))
		eHomePhone			= Null2Space(rs("Phone"))
		eMobilePhone		= Null2Space(rs("Mobile"))
		eSkype				= Null2Space(rs("Skype"))
		eHarvestID			= Null2Space(rs("HarvestID"))

		eCreatedBy			= Null2Space(rs("CreatedBy"))
		eCreatedDate		= Null2Space(rs("CreatedDate"))
		eEditedBy			= Null2Space(rs("EditedBy"))
		eEditedDate			= Null2Space(rs("EditedDate"))
		
        rs.close
	End If
		
%>

<html>
<head>
  <title><%=Title%></title>

  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
  
  <link href="/ha_Backoffice/includes/ha_BackOfficePopup.css" type="text/css" rel="stylesheet">
  <script language="javascript" type="text/javascript" src="/includes/haFunctions.js"></script>
</head>

<body onLoad="SetScreen(650,650);">
  <form method="Post" action="ha_Employee_post.asp?ID=<%=ID%>" name="frm">

	  <span class="popupHeading"><%=Title%></span>
	  
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">	  
  <tr>
    <td colspan="2">
	
      <table width="95%" border="0" cellspacing="2" cellpadding="1" align="center">

		<tr><td>&nbsp;</td></tr>
			
		<tr>
		  <td width="35%" align="right">
			First Name:
		  </td>
		  <td width="65%">
			<input type="text" name="FirstName" id="FirstName"
				size="20" value="<%=eFirstName%>">
		  </td>		 
		</tr>

		<tr>
		  <td align="right">
			Last Name:
		  </td>
		  <td>
			<input type="text" name="LastName" id="LastName"
				size="20" value="<%=eLastName%>">
		  </td>		 
		</tr>
		
		<tr>
		  <td align="right">
			E-mail:
		  </td>
		  <td>
			<input type="text" name="Email" id="Email"
				size="20" value="<%=eEmail%>">
		  </td>		 
		</tr>
		
		<tr>
		  <td align="right">
			Date of Birth:
		  </td>
		  <td>
			<input type="text" name="DOB" id="DOB"
				size="10" value="<%=iif(HasPermission("Read", 9),eDOB,Bullets8)%>"
				<%=iif(HasPermission("Edit", 9),""," Disabled")%>>
			<%=iif(HasPermission("Edit", 9),PopupDate("DOB"),"")%>
		  </td>		 
		</tr>
		
		<tr>
		  <td align="right" style="height: 22px;">
			Birthday:
		  </td>
		  <td>
			<%=eBirthday%>
		  </td>		 
		</tr>
		
		<tr>
		  <td align="right">
			Address:
		  </td>
		  <td>
			<input type="text" name="AddressLine1" id="AddressLine1"
				size="20" value="<%=eAddressLine1%>">
		  </td>		 
		</tr>
		
		<tr>
		  <td align="right">
			Address 2:
		  </td>
		  <td>
			<input type="text" name="AddressLine2" id="AddressLine2"
				size="20" value="<%=eAddressLine2%>">
		  </td>		 
		</tr>
		
		<tr>
		  <td align="right">
			City:
		  </td>
		  <td>
			<input type="text" name="City" id="City"
				size="20" value="<%=eCity%>">
		  </td>		 
		</tr>
		
		<tr>
		  <td align="right">
			State:
		  </td>
		  <td>
			<input type="text" name="State" id="State"
				size="20" value="<%=eState%>">
		  </td>		 
		</tr>

		<tr>
		  <td align="right">
			Postal Code:
		  </td>
		  <td>
			<input type="text" name="PostalCode" id="PostalCode"
				size="20" value="<%=ePostalCode%>">
		  </td>		 
		</tr>

		<tr>
		  <td align="right">
			Work Location:
		  </td>
		  <td>
			<input type="text" name="WorkLocation" id="WorkLocation"
				size="20" value="<%=eWorkLocation%>">
		  </td>		 
		</tr>
		
		<tr>
		  <td align="right">
			Home Phone:
		  </td>
		  <td>
			<input type="text" name="HomePhone" id="HomePhone"
				size="20" value="<%=eHomePhone%>">
		  </td>		 
		</tr>
		
		<tr>
		  <td align="right">
			Mobile Phone:
		  </td>
		  <td>
			<input type="text" name="MobilePhone" id="MobilePhone"
				size="20" value="<%=eMobilePhone%>">
		  </td>		 
		</tr>

		<tr>
		  <td align="right">
			Skype:
		  </td>
		  <td>
			<input type="text" name="Skype" id="Skype"
				size="20" value="<%=eSkype%>">
		  </td>		 
		</tr>
		
		<tr>
		  <td align="right">
			Harvest ID:
		  </td>
		  <td>
			<input type="text" name="HarvestID" id="HarvestID"
				size="20" value="<%=eHarvestID%>">
		  </td>		 
		</tr>
		
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
</body>
</html>
