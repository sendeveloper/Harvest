<!DOCTYPE html>
<!--#include virtual="includes/connection.asp"-->
<!--#include virtual="ha_BackOffice/includes/Config.asp"-->

<%
	ColorTab = 10
	PageHeading = "Online Payment Process"
%>

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<title>Harvest American Backoffice - <%=PageHeading%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    
	<link href="<%=strPathIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet" />
	<link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>menu/dropdowntabfiles/halfmoontabs.css" />
	
	<script type="text/javascript" src="<%=strPathIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>
	<script type="text/javascript" src="<%=strPathBOIncludes%>haFunctions.js" language="javascript"></script>
  
    <style type="text/css">
		td
			{
			vertical-align:	top;
			}
			
		th
			{
			border-bottom:	1px solid black;
			}
			
		div.members
			{
			border:			1px solid gray;
			height:			7em;
			}
    </style>
</head>

<body class="gray_desktop">
  <table class="MainBodyFixed">	
    <tr valign="top">
      <td>
        <!--#include virtual="ha_Backoffice/includes/bodyparts/ha_PageHeading.inc"-->
        <table class="MainTable">
	
		  <tr><td>&nbsp;</td></tr>
	
		  <tr>
			<td align="center">
			  <table width="95%" cellpadding="2" cellspacing="2">
				<tr>
				  <td>
					A Typical Website Flowchart
				  </td>
				</tr>

				<tr><td>&nbsp;</td></tr>
		  
		
				<tr>
                	<td>
                    	<b>1)- &nbsp;&nbsp;<b><a href="javascript:clickPopupImage(1560);" class="TopRightLink">Website Flowchart</a></b>&nbsp
                    </td>
                </tr>

                
				<tr>
				  <td colspan="4">
					<hr>
				  </td>
				</tr>
		  
			  </table>
			</td>
		  </tr>
		  
		</table>
      </td>
    </tr>
	
  </table>
  <!--#include virtual="ha_BackOffice/includes/BodyParts/ha_PageTail.inc"-->
</body>
</html>

