<%
  Function FormatPhone(aPhone)
    'Set local variables
    Dim strPhone : strPhone = aPhone
    Dim strHolder
    'Response.write Len(aPhone)
    'If phone is invalid over length
    If Len(aPhone) = 7 Then

      'If phone is without area code
      strHolder = Left(strPhone,3) & "-" & Right(strPhone,4)
      strPhone = strHolder

    'Else if phone as a length of 10
    ElseIf Len(aPhone) = 10 Then

      'If phone has area code
      strHolder = Left(strPhone,3) & "-"
      strHolder = strHolder & Mid(strPhone,4,3) & "-"
      strHolder = strHolder & Right(strPhone,4)
      strPhone = strHolder

    End If

    'Return phone string
    FormatPhone = strPhone

  End Function

  Function RemoveFormatting(aString)
    
    'Declare local variables
    Dim strString, character

    'For every element in aPhone
    For index = 0 To Len(aString)
      character = right(left(aString, index), 1)
      
      'If the character in the phone is not a number or the character is punctuated
      If InStr(character, "(") Or InStr(character, ")") Or InStr(character, "-") Then
        strString = strString & ""
      Else
        strString = strString & character
      End If
    Next
    'response.write strString
    'Return phone string
    RemoveFormatting = FormatPhone(strString)
  End Function
%>