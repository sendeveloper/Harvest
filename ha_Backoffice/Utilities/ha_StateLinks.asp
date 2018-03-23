<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">



<html>
<head>

  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/Config.asp"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->
  <!--#include virtual="ha_Backoffice/Utilities/connection_Barley.asp"-->

<%
	ColorTab = 0
	PageHeading = "State Gambling Regulation Links Edit Page"

    dim strFilter
    dim strColor
    dim strHeading
    dim LineCount
    dim recordCount
    dim findSQL
    dim abbr(100)
    dim desc(100)



    SQL="SELECT * " & _
	"FROM ni_StateRegulations " & _
        "ORDER BY [Description]"

    set rs=server.createObject("ADODB.Recordset")
    rs.open SQL,objcon,2,3

    if not RS.EOF then
	do while not rs.eof
	    recordCount = recordCount + 1
	    abbr(recordCount) = rs("Abbreviation")
	    desc(recordCount) = rs("Description")
            If isnull(rs("URL")) or rs("URL") = "" then
                desc(recordCount) = "<i>" & desc(recordCount) & "</i>"
            End If
	    rs.movenext
	loop
    end if		    
    rs.close
    set rs = nothing
		
%>



    <title>Harvest American Backoffice - <%=PageHeading%></title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

    <link href="<%=strPathIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet">
    <link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>menu/dropdowntabfiles/halfmoontabs.css">

    <script type="text/javascript" src="<%=strPathIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>

    <style type="text/css">
    </style>

    <script language="javascript" type = "text/javascript">
    </script>
</head>

<body class="gray_desktop">
  <table align="center" cellspacing="0" cellpadding="0" class="MainBody">
    <tr>
      <td>
	<!--#include virtual="ha_Backoffice/includes/bodyparts/ha_PageHeading.inc"-->
      </td>
    </tr>

    <tr>
		<td align="center">
		This page allows you to edit and update the links to individual state's and province's gaming regulations websites found 
		on <a href="http://raffleticketsoftware.com/help/faq.asp">http://raffleticketsoftware.com/help/faq.asp</a><br>
		&nbsp
			<table width='90%'>
				<table width="100%" border="0" cellspacing="5" cellpadding="0" align="center">


					<%
						For i = 1 to 13
							If i mod 5 = 0 then
							'strColor = "#C8FFC8"
							Else
							strColor = "#FFFFFF"
							End If
					%>
						
						<tr bgcolor=<%=strColor%>>
						  <td width="5%">
									&nbsp;
								  </td>
					<%
							For ii = 0 to 4
					%>

						  <td width="19%">
						  <a href="javascript:window.open('boStateLinkEdit.asp?State=<%=abbr(i + (ii * 13))%>',
							'','scrollbars=no,fullscreen=no,resizable=no, 
							height=250,width=750,left=10,top=100');void(0)">
									<FONT Size="1"><%=desc(i + (ii * 13))%></FONT></a>
						  </td>

					<%
							Next
					%>

						</tr>

					<%
							Next
					%>

				</table>
			</td>
		</tr>
<%
	For i = 1 to 2
		Response.Write "<tr><td>&nbsp;</td></tr>"
	Next
%>
  </table>
  </table>
  <!--#include virtual="ha_BackOffice/includes/BodyParts/ha_PageTail.inc"-->
  
</body>
</html>

