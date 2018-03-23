<!--#include virtual="ha_backoffice/includes/Config.asp"-->
<!--#include virtual="includes/connection.asp"-->

<%
	ColorTab = 2
	PageHeading = "Telephone Numbers"
%>

<html>
<head>
    <title>Harvest American Backoffice - Telephone Numbers</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />

    <link href="<%=strPathBOIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet" />
    <link rel="stylesheet" type="text/css" href="<%=strPathBOIncludes%>menu/dropdowntabfiles/halfmoontabs.css" />

    <script type="text/javascript" src="<%=strPathBOIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>

    <style type="text/css">
		th
		{
		  text-align :       center;
          vertical-align :   top;
          background-color : #008000;
          font-weight :      bold;
          color :            white;
		}

		td.Heading1
		{
		  font-size:		14px;
		  text-align:		center;
		  font-weight:      bold;
		  border-bottom:	1px solid black;
		}

    </style>

</head>

<body class="gray_desktop">
  <table align="center" cellspacing="0" cellpadding="0" class="MainBody">
    <tr>
	  <td>

		<!--#include virtual="ha_Backoffice/includes/bodyparts/ha_PageHeading.inc"-->
	
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
	
		  <tr><td>&nbsp;</td></tr>
		
		  <tr>
		    <td align="center">
			  <table width="75%" border="0" cellspacing="5" cellpadding="5">

				<tr valign="top">
				  <td class="Heading1" colspan="6">Harvest American, Inc.</td>
				</tr>

				<tr>
				  <th>Purpose</th>
				  <th>Number</th>
				  <th>Alt Number</th>
				  <th>If Busy</th>
				  <th>If No Answer</th>
				  <th>Vendor</th>
				</tr>
				
				<tr style="background-color: #B4D7BF;">
				  <td>Line 1</td>
				  <td>315-245-1122</td>
				  <td>888-680-2813</td>
				  <td>Line 2</td>
				  <td>Line 2</td>
				  <td>Vonage</td>
				</tr>
				
				<tr>
				  <td>Line 2</td>
				  <td>315-533-1298</td>
				  <td>877-295-8524<br />0207-096-1969 (UK)</td>
				  <td>&nbsp;</td>
				  <td>Vonage Voicemail</td>
				  <td>Vonage</td>
				</tr>

				<tr style="background-color: #B4D7BF;">
				  <td>Fax</td>
				  <td>866-488-6481</td>
				  <td>&nbsp;</td>
				  <td>&nbsp;</td>
				  <td>&nbsp;</td>
				  <td>&nbsp;</td>
				</tr>

				<tr><td>&nbsp;</td></tr>

				<tr valign="top">
				  <td class="Heading1" colspan="6">Zip2Tax, LLC</td>
				</tr>

				<tr>
				  <th>Purpose</th>
				  <th>Number</th>
				  <th>Alt Number</th>
				  <th>If Busy</th>
				  <th>If No Answer</th>
				  <th>Vendor</th>
				</tr>
				
				<tr style="background-color: #B4D7BF;">
				  <td>Line 1</td>
				  <td>307-462-0195</td>
				  <td>866-492-8494</td>
				  <td>Line 2</td>
				  <td>Vonage Voicemail</td>
				  <td>Vonage</td>
				</tr>
				
				<tr>
				  <td>Line 2</td>
				  <td>307-316-8709</td>
				  <td>&nbsp;</td>
				  <td>Line 3</td>
				  <td>Vonage Voicemail</td>
				  <td>Vonage</td>
				</tr>

				<tr style="background-color: #B4D7BF;">
				  <td>Line 3</td>
				  <td>307-316-7275</td>
				  <td>&nbsp;</td>
				  <td>&nbsp;</td>
				  <td>Vonage Voicemail</td>
				  <td>Vonage</td>
				</tr>
				
				<tr>
				  <td>Fax</td>
				  <td>866-488-6481</td>
				  <td>&nbsp;</td>
				  <td>&nbsp;</td>
				  <td>&nbsp;</td>
				  <td>&nbsp;</td>
				</tr>

				<tr><td>&nbsp;</td></tr>
				</table>
                <table width="75%" border="0" cellspacing="5" cellpadding="5">
				<tr valign="top">
				  <td class="Heading1" colspan="4">Cell Phones</td>
				</tr>

				<tr>
				  <th>Assignee</th>
				  <th>Number</th>
                  <th>Vendor / Phone Type</th>
				  <th>Contract Expiration Date</th>
				</tr>

                <tr style="background-color: #B4D7BF;">
				  <td>Chris</td>
				  <td>315-335-8084</td>
                  <td>Verizon / iPhone 6</td>
				  <td></td>
				</tr>

				<tr>
				  <td>Dave</td>
				  <td>315-709-7867</td>
				  <td>Verizon / LG G Pad X8.3</td>
				  <td></td>
				</tr>
				
				<tr style="background-color: #B4D7BF;">
				  <td>Dave</td>
				  <td>315-709-7145</td>
				  <td>Verizon / Samsung Galaxy Note8</td>
				  <td></td>
				</tr>

				<tr>
				  <td>Teresa F.</td>
				  <td>315-832-8623</td>
				  <td>Verizon / Galaxy Note 4</td>
				  <td></td>
				</tr>
				
				<tr><td>&nbsp;</td></tr>

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

