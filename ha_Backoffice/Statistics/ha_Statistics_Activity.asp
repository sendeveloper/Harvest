<!DOCTYPE html>
<html>
<head>
  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/Config.asp"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsDates.inc"-->
  

<%
	ColorTab = 6
	PageHeading = "Activity"

    dim HeadingDateText
    dim StartDate
    dim EndDate
    dim eDate

    if isnull(Request("StartDate")) or Trim(Request("StartDate")) = "" then
      StartDate = date()
    else
      StartDate = Request("StartDate")
    end if
    eDate = StartDate

    if isnull(Request("EndDate")) or Trim(Request("EndDate")) = "" then
      EndDate = date()
    else
      EndDate = Request("EndDate")
    end if

    if isnull(Request("reqDomain")) or Trim(Request("reqDomain")) = "" then
      reqDomain = ""
    else
      reqDomain = Request("reqDomain")
    end if
	
    if isnull(Request("reqPage")) or Trim(Request("reqPage")) = "" then
      reqPage = ""
    else
      reqPage = Request("reqPage")
    end if
	
	If StartDate = EndDate Then
		HeadingDateText = StartDate
	Else
		HeadingDateText = "From " & StartDate & " to " & EndDate
	End If
	
    Dim rs: Set rs = Server.CreateObject("ADODB.Recordset")
	SQL = "ha_Activity_Statistics_DayType_list('" & StartDate & "', '" & EndDate & "', '" & reqDomain & "', '" & reqPage & "')"
	'response.write reqDomain & "|<br>"
	'response.write reqPage & "|<br>"
	'response.write StartDate & "|<br>"
	'response.write EndDate & "|<br>"
	'response.write SQL & "<br>"
    rs.Open SQL, connLocal, 3, 3, 4		  
	
%>
  
  <title>Harvest American Backoffice - <%=PageHeading%></title>
  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

  <link href="<%=strPathIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet">
  <link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>menu/dropdowntabfiles/halfmoontabs.css">

  <script type="text/javascript" src="<%=strPathIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>

  <style type="text/css">
  
	td.subHead
		{
		font-weight: bold;
		border-bottom: 1px solid black;
		}
		
    .buttonFilter
		{
        font-weight: bold;
        font-size: 8pt;
        font-family: Verdana, Arial, Helvetica, sans-serif;
        color: #246B44;
		width: 50px;
		display: inline-block;
        padding: 1px 2px;
        background-color: #D8F0E3;
        border-top: 1px solid #C0C0C0;
        border-right: 1px solid black;
        border-bottom: 1px solid black;
        border-left: 1px solid #C0C0C0;
        text-align: center;
        text-decoration: none;
		margin-right: 2px;
		}
    
    .buttonFilter:hover
		{
        color: white;
        border-color: black #C0C0C0 #C0C0C0 black;
		}
	  
	.button90
		{
		width: 90px;
		}
  </style>

  <script language="javascript" type = "text/javascript">
  
		function formLoad()
			{
            document.getElementById('StartDate').value = '<%=StartDate%>';
            document.getElementById('EndDate').value = '<%=EndDate%>';
			}
  
        function collapseDomain()
                {
                document.getElementById('reqDomain').value = '';
                document.getElementById('reqPage').value = '';
                document.frm.submit();
                }

        function expandDomain(d)
                {
				if (d == '')
					{
					d = '(blank)'
					}
                document.getElementById('reqDomain').value = d;
                document.frm.submit();
                }
				
        function collapsePage()
                {
                document.getElementById('reqPage').value = '';
                document.frm.submit();
                }

        function expandPage(p)
                {
				if (p == '')
					{
					p = '(blank)'
					}				
                document.getElementById('reqPage').value = p;
                document.frm.submit();
                }
				
		function changeDates(StartDate, EndDate) 
			{
			if (EndDate == '') 
				{
				EndDate = StartDate;
				}
			document.frm.StartDate.value = StartDate;
			document.frm.EndDate.value = EndDate;
			clickSubmit();
			}

		function checkStartDate() 
			{
			var d = document.frm.StartDate;
			if (d.value.length != 0) 
				{
				if (isDate(d.value)==false) 
					{
					d.focus();
					return false;
					}
				return true;
				}
				alert('Please enter a start date');
				d.focus();
				return false;
			}

		function checkEndDate() 
			{
			var d = document.frm.EndDate;
			if (d.value.length != 0) 
				{
				if (isDate(d.value)==false) 
					{
					d.focus();
					return false;
					}
				return true;
				}
				alert('Please enter an end date');
				d.focus();
				return false;
			}

		function clickSubmit() 
			{
			if (checkStartDate() && checkEndDate()) 
				{
				document.frm.submit();
				}
			}
  </script>
</head>

<body class="gray_desktop" onLoad="formLoad();">
  
  <table cellspacing="0" cellpadding="0" class="MainBodyFixed" align="center">
    <tr>
	  <td>
	  
		<!--#include virtual="ha_Backoffice/includes/bodyparts/ha_PageHeading.inc"-->
	  
		<table style="width: 1150px; margin: 0 auto;" border="0" cellspacing="0" cellpadding="0">
		  <tr>
			<td width="60%" valign="top">
	  
			  <table width="780" height="350" cellspacing="0" cellpadding="0" style="border: 1px solid black;">
				<tr>
				  <td width="100%" valign="top">
				    <table width="100%" cellspacing="2" cellpadding="2">
					  <tr>
						<td style="text-align: center; font-weight: bold;" colspan="4">
						  <%=HeadingDateText%>
						</td>
					  </tr>
					  <tr>
						<td width="628" class="subHead">
						  Domain/Page/Operation
						</td>
						<td width="50" class="subHead" align="center">
						  Hits
						</td>
						<td width="50" class="subHead" align="center">
						  Uniques
						</td>
						<td>
						  &nbsp;
						</td>
					  </tr>
					</table>
					
				    <div style="overflow-y: scroll; height: 700px;">
					  <table width="100%" cellspacing="2" cellpadding="2">
					
<%
	If rs.eof Then
		Response.Write "No Records"
	Else
		While not rs.eof
			pmImage = ""
			
			recordCount = recordCount + 1
			If recordCount mod 3 = 0 Then
				Underscore = "style='border-bottom: 2px solid #D8F0E3;' "
			Else
				Underscore = ""
			End If
			
			Select Case rs("statLevel")
			Case 1
				If rs("rowsPages") > 1 Then
					pmImage = "plus.png"
					onClick = "expandDomain('" & rs("statDomain") & "')"
				End If
				If rs("statDomain") = reqDomain Then
					pmImage = "minus.png"
					onClick = "collapseDomain()"
				End If
			Case 2
				If rs("rowsOperations") > 1 Then
					pmImage = "plus.png"
					onClick = "expandPage('" & rs("statPage") & "')"
				End If
				If rs("statPage") = reqPage Then
					pmImage = "minus.png"
					onClick = "collapsePage()"
				End If
			End Select
			
			If pmImage > "" Then
				pmImage = "<img src='/ha_backoffice/images/" & pmImage & "' onClick=""JavaScript:" & onClick & ";"">"
			Else
				pmImage = "&nbsp;"
			End If
			
			Select Case rs("statLevel")
			Case 1
%>
						<tr>
						  <td>
						    <%=pmImage%>
						  </td>
						  <td <%=Underscore%>colspan="2">
							<%=rs("statDomain")%>
						  </td>
						  <td <%=Underscore%>align="right">
						    <%=rs("statCountDomain")%>
						  </td>
						  <td <%=Underscore%>>
						    &nbsp;
						  </td>
						</tr>
<%
			Case 2
%>
						<tr>
						  <td>
						    &nbsp;
						  </td>
						  <td>
						    <%=pmImage%>
						  </td>
						  <td <%=Underscore%>>
							<%=rs("statPage")%>
						  </td>
						  <td <%=Underscore%>align="right">
						    <%=rs("statCountPages")%>
						  </td>
						  <td <%=Underscore%>>
						    &nbsp;
						  </td>
						</tr>
<%
			Case Else
%>
						<tr>
						  <td>
						    &nbsp;
						  </td>
						  <td>
						    &nbsp;
						  </td>
						  <td <%=Underscore%>>
							- <%=rs("statOperation")%>
						  </td>
						  <td <%=Underscore%>align="right">
						    <%=rs("statCount")%>
						  </td>
						  <td <%=Underscore%>align="right">
						    <%=rs("statUniques")%>
						  </td>
						</tr>
<%
			End Select
			rs.movenext
		Wend
		rs.Close
	End If
%>
						<tr>
						  <td width="8">&nbsp;</td>
						  <td width="8">&nbsp;</td>
						  <td width="600">&nbsp;</td>
						  <td width="50">&nbsp;</td>
						  <td width="50">&nbsp;</td>
						  <td>&nbsp;</td>
						</tr>
						
					  </table>
					</div>
				  </td>
				</tr>
			  </table>
			  
			</td>

			<td width="40%" align="right" vAlign="top">
			
              <form method="post" action="ha_Statistics_Activity.asp" name="frm" style="margin: 0;">
			    <input type="hidden" id="reqDomain" name="reqDomain" value="<%=reqDomain%>" />
			    <input type="hidden" id="reqPage" name="reqPage" value="<%=reqPage%>" />

              <table width='350' height='300' border='0' cellspacing='0' cellpadding='0' style="background-color: #ECF7F1;">
                <tr>
                  <td valign='top'>
				    <span style='padding: 10px 0 0 5px; color: #246B44; font-size: 11pt; font-weight: bold; display: block;'>Date Filters</span><br>
                    <table width='100%' border='0' cellspacing='5' cellpadding='5'>
                      <tr>
                        <td width='25%' align='right'>
                          Start&nbsp;Date
                        </td>
                        <td width='45%' style="white-space: nowrap;">
                          <input type="Text" id="StartDate" name="StartDate" value="<%=d1%>" size="10" style="width: 10em;">
                          <a href="javascript:show_calendar('document.frm.StartDate', document.frm.StartDate.value);">
                            <img src="/ha_Backoffice/includes/dates/cal.gif" width="16" height="16"
                                 border="0" alt="Calendar"></a>
                        </td>
                      </tr>
                      
                      <tr>
                        <td align='right'>
                          End&nbsp;Date
                        </td>
                        <td style="white-space: nowrap;">
                          <input type="Text" id="EndDate" name="EndDate" value="<%=d2%>" size="10" style="width: 10em;">
                          <a href="javascript:show_calendar('document.frm.EndDate', document.frm.EndDate.value);">
                            <img src="/ha_Backoffice/includes/dates/cal.gif" width="16" height="16"
                                 border="0" alt="Calendar"></a>
                        </td>
                        <td width='25%' valign='bottom'>
                          <a href="javascript:clickSubmit();" class="buttonFilter">Go</a><br>
                        </td>
                      </tr>
                      
                      <tr>
                        <td colspan='3'>
                          <hr>
                        </td>
                      </tr>
                      
                    </table>
                    
                    <table width='100%' border='0' cellspacing='5' cellpadding='5'>
                      <tr valign='top'>
                        <td align='right'>
                          <a href="javascript:changeDates('<%=DateAdd("d",-1,eDate)%>', '');"
                             class="buttonFilter button90">&lt; Previous</a>
                        </td>
                        <td align='center'>
                          <a href="javascript:changeDates('<%=date()%>', '');" class="buttonFilter button90">
                            Today</a>
                        </td>
                        <td>
                          <a href="javascript:changeDates('<%=DateAdd("d",1,eDate)%>', '');"
                             class="buttonFilter button90">Next &gt;</a>
                        </td>
                      </tr>
                    </table>
                    
                    <table width='100%' border='0' cellspacing='5' cellpadding='5'>
                      <tr valign='top'>
                        <td align="right"></td>
                        <td align='center'>
                          <a href="javascript:changeDates('<%=DateAdd("d", -(Weekday(date) - 1), date())%>','<%=DateAdd("d", 7 - (Weekday(date())) - 1, date())%>');"
                             class="buttonFilter button90">This Week</a>
                        </td>
                        <td align="left">
                        </td>
                      </tr>
                    </table>
                    
                    <table width='100%' border='0' cellspacing='5' cellpadding='5'>
                      <tr>
                        <td align='right'>
<%
    sMonth = Month(DateAdd("m",-1,eDate)) & "/1/" & Year(DateAdd("m",-1,eDate))
    eMonth = DateAdd("d",-1,DateAdd("m",1,sMonth))
%>
                          <a href="javascript:changeDates('<%=sMonth%>', '<%=eMonth%>');"
                             class="buttonFilter button90">&lt; Previous</a>
                        </td>
                        <td align='center'>
<%
    sMonth = Month(date()) & "/1/" & Year(date())
    eMonth = DateAdd("d",-1,DateAdd("m",1,sMonth))
%>
                          <a href="javascript:changeDates('<%=sMonth%>', '<%=eMonth%>');" class="buttonFilter button90">This Month</a>
                        </td>
                        <td>
<%
    sMonth = Month(DateAdd("m",1,eDate)) & "/1/" & Year(DateAdd("m",1,eDate))
    eMonth = DateAdd("d",-1,DateAdd("m",1,sMonth))
%>
                          <a href="javascript:changeDates('<%=sMonth%>', '<%=eMonth%>');"
                             class="buttonFilter button90">Next &gt;</a>
                        </td>
                      </tr>
                    </table>

                    <table width='100%' border='0' cellspacing='5' cellpadding='5'>
                      <tr>
                        <td align='right'>
<%
    sMonth = "1/1/" & Year(DateAdd("yyyy",-1,eDate))
    eMonth = DateAdd("d",-1,DateAdd("yyyy",1,sMonth))
%>
                          <a href="javascript:changeDates('<%=sMonth%>', '<%=eMonth%>');"
                             class="buttonFilter button90">&lt; Previous</a>
                        </td>
                        <td align='center'>
<%
    sMonth = "1/1/" & Year(date())
    eMonth = DateAdd("d",-1,DateAdd("yyyy",1,sMonth))
%>
                          <a href="javascript:changeDates('<%=sMonth%>', '<%=eMonth%>');"
                             class="buttonFilter button90">This Year</a>
                       </td>
                       <td>
<%
    sMonth = "1/1/" & Year(DateAdd("yyyy",1,eDate))
    eMonth = DateAdd("d",-1,DateAdd("yyyy",1,sMonth))
%>
                          <a href="javascript:changeDates('<%=sMonth%>', '<%=eMonth%>');"
                             class="buttonFilter button90">Next &gt;</a>
                        </td>
                      </tr>
                    </table>
					
                  </td>
                </tr>
              </table>
              </form>
			
			</td>
		  </tr>
		  
		  <tr><td>&nbsp;</td></tr>
		</table>
		
	  </td>
	</tr>
  </table>
  
  <!--#include virtual="ha_BackOffice/includes/BodyParts/ha_PageTail.inc"-->
  
</body>
</html>

