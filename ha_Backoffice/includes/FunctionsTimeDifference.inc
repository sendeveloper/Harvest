<%
    Function TimeDifference(x, y)
      Dim sec, min, hour, day, mon
	  
	  if isnull(x) then x = now
	  if isnull(y) then y = now

      sec = DateDiff("s", CDate(x), CDate(y))
      min = DateDiff("n", CDate(x), CDate(y))
      hour = DateDiff("h", CDate(x), CDate(y))
      day = DateDiff("d", CDate(x), CDate(y))
      mon = DateDiff("m", CDate(x), CDate(y))

      If mon > 0 Then
        TimeDifference = mon & " Month(s) Ago"
      ElseIf day > 0 Then
        TimeDifference = day & " Day(s) Ago"
      ElseIf hour > 0 Then
        TimeDifference = hour & " Hour(s) Ago"
      ElseIf min > 0 Then
		If min = 1 Then
			TimeDifference = min & " Minute Ago"
		Else
			TimeDifference = min & " Minutes Ago"
		End If
      ElseIf sec > 0 Then
        TimeDifference = sec & " Seconds Ago"
      Else
        TimeDifference = "Just Now"
      End If
    End Function
%>