<%

Function ConvertCurrencyToEnglish(ByVal strAmount)
         Dim Temp
         Dim Dollars, Cents
         Dim DecimalPlace, Count

         ReDim Place(9) 
         Place(2) = " Thousand "
         Place(3) = " Million "
         Place(4) = " Billion "
         Place(5) = " Trillion "

         ' Convert strAmount to a string, trimming extra spaces.
         strAmount = Trim(cStr(strAmount))

         ' Find decimal place.
         DecimalPlace = InStr(strAmount, ".")

         ' If we find decimal place...
         If DecimalPlace > 0 Then
            ' Convert cents
            Temp = Left(Mid(strAmount, DecimalPlace + 1) & "00", 2)
            Cents = ConvertTens(Temp)

            ' Strip off cents from remainder to convert.
            strAmount = Trim(Left(strAmount, DecimalPlace - 1))
         End If

         Count = 1
         Do While strAmount <> ""
            ' Convert last 3 digits of strAmount to English dollars.
            Temp = ConvertHundreds(Right(strAmount, 3))
            If Temp <> "" Then Dollars = Temp & Place(Count) & Dollars
            If Len(strAmount) > 3 Then
               ' Remove last 3 converted digits from strAmount.
               strAmount = Left(strAmount, Len(strAmount) - 3)
            Else
               strAmount = ""
            End If
            Count = Count + 1
         Loop

         ' Clean up dollars.
         Select Case Dollars
            Case ""
               Dollars = "No Dollars"
            Case "One"
               Dollars = "One Dollar"
            Case Else
               Dollars = Dollars & " Dollars"
         End Select

         ' Clean up cents.
         Select Case Cents
            Case ""
               'Cents = " And No Cents"
            Case "One"
               Cents = " And One Cent"
            Case Else
               Cents = " And " & Cents & " Cents"
         End Select

          ConvertCurrencyToEnglish = Dollars & Cents
      End Function



     Private Function ConvertHundreds(ByVal strAmount)
         Dim Result 

         ' Exit if there is nothing to convert.
		 If not isNumeric(strAmount) Then Exit Function
         If cInt(strAmount) = 0 Then Exit Function

         ' Append leading zeros to number.
         strAmount = Right("000" & strAmount, 3)

         ' Do we have a hundreds place digit to convert?
         If Left(strAmount, 1) <> "0" Then
            Result = ConvertDigit(Left(strAmount, 1)) & " Hundred "
         End If

         ' Do we have a tens place digit to convert?
         If Mid(strAmount, 2, 1) <> "0" Then
            Result = Result & ConvertTens(Mid(strAmount, 2))
         Else
            ' If not, then convert the ones place digit.
            Result = Result & ConvertDigit(Mid(strAmount, 3))
         End If

         ConvertHundreds = Trim(Result)
      End Function



      Private Function ConvertTens(ByVal MyTens)
         Dim Result 

		 'Response.write MyTens & "<br>"
		 'exit function
		 
         ' Exit if there is nothing to convert.
		 if MyTens = "" Then Exit Function
		 If not isNumeric(MyTens) Then Exit Function
		 
         ' Is value between 10 and 19?
         If cInt(Left(MyTens, 1)) = 1 Then
            Select Case cInt(MyTens)
               Case 10: Result = "Ten"
               Case 11: Result = "Eleven"
               Case 12: Result = "Twelve"
               Case 13: Result = "Thirteen"
               Case 14: Result = "Fourteen"
               Case 15: Result = "Fifteen"
               Case 16: Result = "Sixteen"
               Case 17: Result = "Seventeen"
               Case 18: Result = "Eighteen"
               Case 19: Result = "Nineteen"
               Case Else
            End Select
         Else
            ' .. otherwise it's between 20 and 99.
            Select Case cInt(Left(MyTens, 1))
               Case 2: Result = "Twenty "
               Case 3: Result = "Thirty "
               Case 4: Result = "Forty "
               Case 5: Result = "Fifty "
               Case 6: Result = "Sixty "
               Case 7: Result = "Seventy "
               Case 8: Result = "Eighty "
               Case 9: Result = "Ninety "
               Case Else
            End Select

            ' Convert ones place digit.
            Result = Result & ConvertDigit(Right(MyTens, 1))
         End If

         ConvertTens = Result
     End Function



      Private Function ConvertDigit(ByVal MyDigit)
         ' Exit if there is nothing to convert.
		 If not isNumeric(MyDigit) Then Exit Function
         If cInt(MyDigit) = 0 Then Exit Function
		 
         Select Case cInt(MyDigit)
            Case 1: ConvertDigit = "One"
            Case 2: ConvertDigit = "Two"
            Case 3: ConvertDigit = "Three"
            Case 4: ConvertDigit = "Four"
            Case 5: ConvertDigit = "Five"
            Case 6: ConvertDigit = "Six"
            Case 7: ConvertDigit = "Seven"
            Case 8: ConvertDigit = "Eight"
            Case 9: ConvertDigit = "Nine"
            Case Else: ConvertDigit = ""
         End Select
      End Function
	  
%>