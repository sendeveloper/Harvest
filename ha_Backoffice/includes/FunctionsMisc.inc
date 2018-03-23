<%
	Function RequestParse(r, d)
		If isnull(r) or r = "" or instr(r,"'") Then
			RequestParse = d
		Else
			RequestParse = r
		End If
	End Function

  Function ColumnHeading(head)
  
    Dim s: s= ""
	s = s & "<table border='0' cellpadding='0' cellspacing='0'>"
    s = s & "  <tr>"
    s = s & "    <td class='head'>" & head & "&nbsp;</td>"
    s = s & "    <td>"
	
	s = s & "      <table cellpadding='0' cellspacing='0'>"
    s = s & "        <tr>"
    s = s & "          <td>"
	
		
    s = s & "      <a href='?OrderBy=" & head & "&ad=ASC' title='Ascending' style='display: block; height: 11px;'><img src='" & strPathImages & "ArrowAscending.jpg' alt='Up'></a>"
    s = s & "      <a href='?OrderBy=" & head & "&ad=DESC' title='Descending' style='display: block; height: 11px;'><img src='" & strPathImages & "ArrowDescending.jpg' alt='Down'></a>"
		
    s = s & "          </td>"
    s = s & "        </tr>"
    s = s & "      </table>"
	
    s = s & "    </td>"
    s = s & "  </tr>"
    s = s & "</table>"
	
	ColumnHeading = s
  End Function
  
  Function TopRightLinks(SkippedColumns, Link1, Link2, Link3, Link4)

	Dim s, a1, a2, a3, a4
	
	a1 = Replace(Link1," ", "")
	a2 = Replace(Link2," ", "")
	a3 = Replace(Link3," ", "")
	a4 = Replace(Link4," ", "")
	
	s = vbCrLf & _
					"<tr valign='top'><td colspan='9' align='right'><table cellspacing='0' cellpadding='0'>" & vbCrLf & _
					"<tr valign='top' style='font-weight: bold;'>" & vbCrLf & _
					"  <td>" & vbCrLf
					
	If Link1 <> "" Then
		s = s & "<a href='javascript:click" & a1 & "();'>" & Link1 & "</a>&emsp;" & vbCrLf
	End If

	If Link2 <> "" Then
		s = s & "<a href='javascript:click" & a2 & "();'>" & Link2 & "</a>&emsp;" & vbCrLf
	End If

	If Link3 <> "" Then
		s = s & "<a href='javascript:click" & a3 & "();'>" & Link3 & "</a>&emsp;" & vbCrLf
	End If
	
	If Link4 <> "" Then
		s = s & "<a href='javascript:click" & a4 & "();'>" & Link4 & "</a>" & vbCrLf
	End If
	
	s = s & "</td>" & vbCrLf & "</tr>" & vbCrLf
	s = s & "</table>" & cvCrLf & "</td>" & vbCrLf & "</tr>" & vbCrLf

	TopRightLinks= s
  End Function
  
  Function TopRightLinksExpanded(Link0, Link1, Link2, Link3, Link4, Link5, Link6, Link7, Link8, Link9)

	Dim s, a0, a1, a2, a3, a4, a5, a6, a7, a8, a9
	
	a0 = Replace(Link0," ", "")
	a1 = Replace(Link1," ", "")
	a2 = Replace(Link2," ", "")
	a3 = Replace(Link3," ", "")
	a4 = Replace(Link4," ", "")
	a5 = Replace(Link5," ", "")
	a6 = Replace(Link6," ", "")
	a7 = Replace(Link7," ", "")
	a8 = Replace(Link8," ", "")
	a9 = Replace(Link9," ", "")
	
	s = vbCrLf & _
					"<table width='100%' cellspacing='0' cellpladding='0'><tr valign='top' style='font-weight: bold;'>" & vbCrLf & _
					"  <td style='text-align: right;'>" & vbCrLf
					
	If Link0 <> "" Then
		s = s & "<a href='javascript:click" & a0 & "();'>" & Link1 & "</a>&emsp;" & vbCrLf
	End If
	
	If Link1 <> "" Then
		s = s & "<a href='javascript:click" & a1 & "();'>" & Link1 & "</a>&emsp;" & vbCrLf
	End If

	If Link2 <> "" Then
		s = s & "<a href='javascript:click" & a2 & "();'>" & Link2 & "</a>&emsp;" & vbCrLf
	End If

	If Link3 <> "" Then
		s = s & "<a href='javascript:click" & a3 & "();'>" & Link3 & "</a>&emsp;" & vbCrLf
	End If
	
	If Link4 <> "" Then
		s = s & "<a href='javascript:click" & a4 & "();'>" & Link4 & "</a>&emsp;" & vbCrLf
	End If

	If Link5 <> "" Then
		s = s & "<a href='javascript:click" & a5 & "();'>" & Link5 & "</a>&emsp;" & vbCrLf
	End If

	If Link6 <> "" Then
		s = s & "<a href='javascript:click" & a6 & "();'>" & Link6 & "</a>&emsp;" & vbCrLf
	End If

	If Link7 <> "" Then
		s = s & "<a href='javascript:click" & a7 & "();'>" & Link7 & "</a>&emsp;" & vbCrLf
	End If

	If Link8 <> "" Then
		s = s & "<a href='javascript:click" & a8 & "();'>" & Link8 & "</a>&emsp;" & vbCrLf
	End If

	If Link9 <> "" Then
		s = s & "<a href='javascript:click" & a9 & "();'>" & Link9 & "</a>" & vbCrLf
	End If
	
	s = s & "</td>" & vbCrLf & "</tr>" & vbCrLf & "</table>" & vbCrLf

	TopRightLinksExpanded= s
  End Function
  
  Function IIf(i,j,k)
    If i Then IIf = j Else IIf = k
  End Function
			  
  Function isNullSpace(value)
    If isNull(value) Then
		isNullSpace = True
	Else
		If value = "" Then
			isNullSpace = True
		Else
			isNullSpace = False
		End If
	End If
  End Function
  
  Function Null2Space(value)
  
	Dim s
	
    If isNull(value) Then
		s = ""
	Else
		s = ltrim(rtrim(value))
	End If
	
	'Removes leading tabs, carriage returns and line feeds
	If s > "" Then
		Do While asc(left(s,1)) = 9 or asc(left(s,1)) = 10 or asc(left(s,1)) = 13
			s = ltrim(mid(s,2))
			If s = "" Then Exit Do
		Loop
	End If
		
	Null2Space = s
	
  End Function
  
  
  Function FixAps(s)
  
	FixAps = Replace(s,"'","''")
	
  End Function

  Function CreditLine(eCreatedBy, eCreatedDate, eEditedBy, eEditedDate)
  	  
	CreditLine = "<span class='popupCreditLeft'>" & CreditInfo("Created", eCreatedBy, eCreatedDate) & "</span>" & _
		"<span class='popupCreditRight'>" & CreditInfo("Edited", eEditedBy, eEditedDate) & "</span>"
  End Function
  
  
  Function CreditInfo(iPrefix, iWho, iWhen)
	Dim s
	
	If iWho > "" Then
		s = "by " & iWho & " "
	End If
	
	If iWhen > "" Then
		s = s & "on " & ConvDate(iWhen, "%a, %b %d, %Y  %I:%M %P")
	End If
	
	If s > "" Then
		s = iPrefix & " " & s
	End If
	
	CreditInfo = s
	
  End Function

Function ConvDate(strDate, strFormat)

   Dim intPosItem
   Dim intHourPart
   Dim strHourPart
   Dim strMinutePart
   Dim strSecondPart
   Dim strAMPM

   If not IsDate(strDate) Then
      ConvDate = strDate
      Exit Function
   End If
	
   intPosItem = Instr(strFormat, "%m")
   Do While intPosItem > 0
      strFormat = Left(strFormat, intPosItem-1) & _
      DatePart("m",strDate) & _
      Right(strFormat, Len(strFormat) - (intPosItem + 1))
      intPosItem = Instr(strFormat, "%m")
   Loop

   intPosItem = Instr(strFormat, "%b")
   Do While intPosItem > 0
      strFormat = Left(strFormat, intPosItem-1) & _
      MonthName(DatePart("m",strDate),True) & _
      Right(strFormat, Len(strFormat) - (intPosItem + 1))
      intPosItem = Instr(strFormat, "%b")
   Loop
	
   intPosItem = Instr(strFormat, "%B")
   Do While intPosItem > 0
      strFormat = Left(strFormat, intPosItem-1) & _
      MonthName(DatePart("m",strDate),False) & _
      Right(strFormat, Len(strFormat) - (intPosItem + 1))
      intPosItem = Instr(strFormat, "%B")
   Loop
	
   intPosItem = Instr(strFormat, "%d")
   Do While intPosItem > 0
      strFormat = Left(strFormat, intPosItem-1) & _
      DatePart("d",strDate) & _
      Right(strFormat, Len(strFormat) - (intPosItem + 1))
      intPosItem = Instr(strFormat, "%d")
   Loop

   intPosItem = Instr(strFormat, "%j")
   Do While intPosItem > 0
      strFormat = Left(strFormat, intPosItem-1) & _
      DatePart("y",strDate) & _
      Right(strFormat, Len(strFormat) - (intPosItem + 1))
      intPosItem = Instr(strFormat, "%j")
   Loop

   intPosItem = Instr(strFormat, "%y")
   Do While intPosItem > 0
      strFormat = Left(strFormat, intPosItem-1) & _
      Right(DatePart("yyyy",strDate),2) & _
      Right(strFormat, Len(strFormat) - (intPosItem + 1))
      intPosItem = Instr(strFormat, "%y")
   Loop

   intPosItem = Instr(strFormat, "%Y")
   Do While intPosItem > 0
      strFormat = Left(strFormat, intPosItem-1) & _
      DatePart("yyyy",strDate) & _
      Right(strFormat, Len(strFormat) - (intPosItem + 1))
      intPosItem = Instr(strFormat, "%Y")
   Loop

   intPosItem = Instr(strFormat, "%w")
   Do While intPosItem > 0
      strFormat = Left(strFormat, intPosItem-1) & _
      DatePart("w",strDate,1) & _
      Right(strFormat, Len(strFormat) - (intPosItem + 1))
      intPosItem = Instr(strFormat, "%w")
   Loop

   intPosItem = Instr(strFormat, "%a")
   Do While intPosItem > 0
      strFormat = Left(strFormat, intPosItem-1) & _
      WeekDayName(DatePart("w",strDate,1),True) & _
      Right(strFormat, Len(strFormat) - (intPosItem + 1))
      intPosItem = Instr(strFormat, "%a")
   Loop

   intPosItem = Instr(strFormat, "%A")
   Do While intPosItem > 0
      strFormat = Left(strFormat, intPosItem-1) & _
      WeekDayName(DatePart("w",strDate,1),False) & _
      Right(strFormat, Len(strFormat) - (intPosItem + 1))
      intPosItem = Instr(strFormat, "%A")
   Loop

   intPosItem = Instr(strFormat, "%I")
   Do While intPosItem > 0
      intHourPart = DatePart("h",strDate) mod 12
      if intHourPart = 0 then intHourPart = 12
      strFormat = Left(strFormat, intPosItem-1) & _
      intHourPart & _
      Right(strFormat, Len(strFormat) - (intPosItem + 1))
      intPosItem = Instr(strFormat, "%I")
   Loop

   intPosItem = Instr(strFormat, "%H")
   Do While intPosItem > 0
      strHourPart = DatePart("h",strDate)
      if strHourPart < 10 Then strHourPart = "0" & strHourPart
      strFormat = Left(strFormat, intPosItem-1) & _
      strHourPart & _
      Right(strFormat, Len(strFormat) - (intPosItem + 1))
      intPosItem = Instr(strFormat, "%H")
   Loop

   intPosItem = Instr(strFormat, "%M")
   Do While intPosItem > 0
      strMinutePart = DatePart("n",strDate)
      if strMinutePart < 10 then strMinutePart = "0" & strMinutePart
      strFormat = Left(strFormat, intPosItem-1) & _
      strMinutePart & _
      Right(strFormat, Len(strFormat) - (intPosItem + 1))
      intPosItem = Instr(strFormat, "%M")
   Loop

   intPosItem = Instr(strFormat, "%S")
   Do While intPosItem > 0
      strSecondPart = DatePart("s",strDate)
      if strSecondPart < 10 then strSecondPart = "0" & strSecondPart
      strFormat = Left(strFormat, intPosItem-1) & _
      strSecondPart & _
      Right(strFormat, Len(strFormat) - (intPosItem + 1))
      intPosItem = Instr(strFormat, "%S")
   Loop

   intPosItem = Instr(strFormat, "%P")
   Do While intPosItem > 0
      if DatePart("h",strDate) >= 12 then
         strAMPM = "PM"
      Else
         strAMPM = "AM"
      End If
      strFormat = Left(strFormat, intPosItem-1) & strAMPM & _
      Right(strFormat, Len(strFormat) - (intPosItem + 1))
      intPosItem = Instr(strFormat, "%P")
   Loop

   intPosItem = Instr(strFormat, "%%")
   Do While intPosItem > 0
      strFormat = Left(strFormat, intPosItem-1) & "%" & _
      Right(strFormat, Len(strFormat) - (intPosItem + 1))
      intPosItem = Instr(strFormat, "%%")
   Loop

   ConvDate = strFormat
	
	'%m Month as a decimal no. 02
	'%b Abbreviated month name Feb
	'%B Full month name February
	'%d Day of the month 23
	'%j Day of the year 54
	'%y Year without century 98
	'%Y Year with century 1998
	'%w Weekday as integer 5 (0 is Sunday)
	'%a Abbreviated day name Fri
	'%A Weekday Name Friday
	'%I Hour in 12 hour format 12
	'%H Hour in 24 hour format 24
	'%M Minute as an integer 01
	'%S Second as an integer 55
	'%P AM/PM Indicator PM
	'%% Actual Percent sign %% 

End Function

%>