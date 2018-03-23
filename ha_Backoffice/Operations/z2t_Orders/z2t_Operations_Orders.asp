<!DOCTYPE html>
<!--#include virtual="includes/connection.asp"-->
<!--#include virtual="ha_BackOffice/includes/Config.asp"-->

<%
	ColorTab = 3
	PageHeading = "Zip2Tax Order Process"
%>

<html>
<head>
	<title>Harvest American Backoffice - <%=PageHeading%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    
	<link href="<%=strPathIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet" />
	<link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>menu/dropdowntabfiles/halfmoontabs.css" />
	<script type="text/javascript" src="<%=strPathIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>
    <script type="text/javascript" src="<%=strPathBOIncludes%>haFunctions.js"></script>

  
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
			height:			7em;
			}
    </style>
	
	<script language='JavaScript'>

	function clickPopupImage(ID)
		{
		var URL = '/ha_BackOffice/Company/DocumentMaintenance/ha_Document_Show.asp' +
			'?PhotoID=' + ID;
			PopupCenter(URL,'','900','500');
		}

    </script>
	
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
		  
		  <tr><td>&nbsp;</td></tr>
		  
  		  <tr>
		    <td colspan="4">
			  <hr>
			</td>
		  </tr>
		  
		  <%=DisplayLine("Zip2Tax Orders",0)%>
		  <%=DisplayLine("Online Lookup",1)%>
		  <%=DisplayLine("Database Interface",1)%>
		  <%=DisplayLine("Tax Tables",1)%>
		  <%=DisplayLine("Initial Purchase",2)%>
		  <%=DisplayLine("Monthly Updates",2)%>
		  
		  
		  <tr>
		    <td colspan="4">
			  <hr>
			  		  <a href="javascript:clickPopupImage(1604);">Zip2Tax Services Available to the Customer</a>&nbsp;&nbsp;

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

<%
	Function DisplayLine(t,i)
		If i < 2 Then
			st = "font-weight: bold;"
		Else
			st = "font-weight: normal;"
		End If
		s = "<tr><td align='left' style='" & st & "'>"
		For ii = 1 to i
			s = s & "&emsp;&emsp;"
		Next
		s = s & t
		s = s & "</td></tr>"
		
		DisplayLine = s & vbCrLf
		
	End Function
%>