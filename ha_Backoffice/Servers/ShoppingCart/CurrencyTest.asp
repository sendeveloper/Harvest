<!--#include file="Currency2Words.asp"-->

<%

	Response.write Results(123.00)
	Response.write Results(123.45)
	Response.write Results(29.95)
	Response.write Results(1)
	Response.write Results(100)
	Response.write Results(1040.80)
	

	Function Results(d)
		Results = ConvertCurrencyToEnglish(d) & " - " & d & "<br>"
	End Function
%>