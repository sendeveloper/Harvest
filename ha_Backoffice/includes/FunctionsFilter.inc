  <style type="text/css">
  
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
	  
  	.filterCaption
		{
		font-size: 8pt;
		color: #888888;
		margin-left: 5px;
		}

  </style>


<%
  Dim CheckBoxes(9,50)
  Dim CheckBoxClass(9,50)
  Dim FilterCategory(5)
  Dim FunctionsFilter
  Dim FiftyZeros
  
  FiftyZeros = "00000000000000000000000000000000000000000000000000"
  
    Dim rsPref: Set rsPref = Server.CreateObject("ADODB.Recordset")

	strSQL = "ha_Employee_Preference_read(" & Session("ha_EmployeeID") & ", '', '" & strPage & "')"
	'response.write strSQL
	rsPref.Open strSQL, connCasper10, 3, 3, 4
	
	If Not rsPref.eof Then
	  CheckBoxes(1,0) = rsPref("Data1")
	  CheckBoxes(2,0) = rsPref("Data2")
	  CheckBoxes(3,0) = rsPref("Data3")
	  CheckBoxes(4,0) = rsPref("Data4")
	  CheckBoxes(5,0) = rsPref("Data5")
	  CheckBoxes(6,0) = rsPref("Data6")
	  CheckBoxes(7,0) = rsPref("Data7")
	  CheckBoxes(8,0) = rsPref("Data8")
	  CheckBoxes(9,0) = rsPref("Data9")
	End If
	
	rsPref.Close
	Set rsPref = Nothing

  Function CheckBoxRead(n, checkClass)
  
	Dim i
	
	strSQL = "ha_Types_list('" & checkClass & "')"
	rs.Open strSQL, connCasper10, 3, 3, 4
		i = 0
		If Not rs.eof Then
			Do While Not rs.eof
				i = i + 1
				CheckBoxes(n,i) = rs("Value")
				If Mid(rs("Description"),2,1) = "|" Then
					CheckBoxClass(n,i) = Left(rs("Description"),1)
				End If
				rs.MoveNext
			Loop
		End If
	rs.Close

  End Function
  
  Function CheckBoxReadProperties(n, checkDescription)
  
	Dim i
	
	strSQL = "SELECT Value FROM ha_Types_repl " & _
		"WHERE Class like '%properties' " & _
		"AND Description = '" & checkDescription &  "' " & _
		"ORDER BY Class"
	rs.Open strSQL, connCasper10, 0, 1
		i = 0
		If Not rs.eof Then
			Do While Not rs.eof
				i = i + 1
				CheckBoxes(n,i) = rs("Value")
				rs.MoveNext
			Loop
		End If
	rs.Close

  End Function
  
  Function CheckBoxReadYear(n, checkTable, checkField)
  
	Dim i
	
	CheckBoxes(n, 1) = "(blank)"
	strSQL = "SELECT Year(" & checkField & ") Value FROM " & checkTable & " " & _
		"WHERE DeletedDate is null " & _
		"GROUP BY Year(" & checkField & ") " & _
		"ORDER BY Year(" & checkField & ") DESC" 
	rs.Open strSQL, connCasper10, 0, 1
		i = 1
		If Not rs.eof Then
			Do While Not rs.eof
				i = i + 1
				CheckBoxes(n,i) = rs("Value")
				rs.MoveNext
			Loop
		End If
	rs.Close

  End Function

  Function CheckBoxReadTable(n, checkTable, checkField)
  
	Dim i:i = 0
	
	'Check for null (blank) records
	If checkTable = "ha_Websites" Then
		'Adding a blank field at the beginning by hand
		i = i + 1
		CheckBoxes(n,i) = "(blank)"
	Else
		strSQL = "SELECT Count(" & checkField & ") c FROM " & checkTable & " " & _
			"WHERE DeletedDate is null " & _
			"AND isnull(" & checkField & ",'') = ''"
		'response.write strSQL & "<br>"
		rs.Open strSQL, connCasper10, 0, 1
		If not rs.eof Then
			'response.write rs("c") & "<br>"
			If rs("c") > 0 Then
				i = i + 1
				CheckBoxes(n,i) = "(blank)"
			End If
		End If
		rs.Close
	End If
	
	'Get the rest
	strSQL = "SELECT " & checkField & " Value FROM " & checkTable & " " & _
		"WHERE DeletedDate is null " & _
		"AND isnull(" & checkField & ",'') > '' " & _
		"GROUP BY " & checkField & " " & _
		"ORDER BY " & checkField
	rs.Open strSQL, connCasper10, 0, 1
	If Not rs.eof Then
		Do While Not rs.eof
			i = i + 1
			CheckBoxes(n,i) = rs("Value")
			rs.MoveNext
		Loop
	End If
	rs.Close

  End Function

  Function CheckBoxReadTableSimple(n, checkTable, checkField)
  
	Dim i
	
	strSQL = "SELECT " & checkField & " Value FROM " & checkTable & " " & _
		"WHERE isnull(" & checkField & ",'') > '' " & _
		"GROUP BY " & checkField & " " & _
		"ORDER BY " & checkField
	rs.Open strSQL, connCasper10, 0, 1
		i = 0
		If Not rs.eof Then
			Do While Not rs.eof
				i = i + 1
				CheckBoxes(n,i) = rs("Value")
				rs.MoveNext
			Loop
		End If
	rs.Close

  End Function

  
  Function CheckBoxFormat(n, checkName)
  
	Dim i, s, Map, LastCheckBoxClass
	
	Map = CheckBoxes(n,0)
	i = 1
	
	s = "<tr><td style='padding: 15px 0 0 10px;' colspan='2'><b>" & CheckName & "</b></td></tr>"
	'After displaying with spaces take it out
	checkName = Replace(checkName," ","")
	
	Do while CheckBoxes(n,i)> ""
		s = s & "<tr>"
		
		If CheckBoxClass(n,i) > "" Then
			s = s & "<td align='right' style='border-right: 1px solid #C0C0C0;'>"
			If CheckBoxClass(n,i) <> LastCheckBoxClass Then
				s = s & "<input type='checkbox'>"
				LastCheckBoxClass = CheckBoxClass(n,i) 
			End If
		Else
			s = s & "<td></td>"
		End If

		s = s & "<td>"
		s = s & "<input type='checkbox' name='new_" & checkName & "' value='" & cstr(i) & "'"
		If mid(Map,i,1) = "1" Then
			s = s & " checked='checked'"
		End IF
		's = s & " style='margin-left: 5px;'>"
		s = s & ">"
		s = s & "<span class='filterCaption'>" & CheckBoxes(n,i) & "</span><br>"
		i = i + 1
		
		s = s & "</td></tr>"
	Loop
	
	CheckBoxFormat = s
	
  End Function
  
  Function FilterCategories(f1, f2, f3, f4, f5)

	Dim s, t, i
	
	FilterCategory(1) = f1
	FilterCategory(2) = f2
	FilterCategory(3) = f3
	FilterCategory(4) = f4
	FilterCategory(5) = f5
	
	s = vbCrLf
	t1 = "    "
	t2 = "        "
	t3 = "            "	
	t4 = "                "	
	
	For i = 1 to 5
		If FilterCategory(i) > "" Then
			s = s & t1 & "var value" & FilterCategory(i) & ";" & vbCrLf
		End If
	Next
	s = s & vbCrLf
	
	s = s & t1 & "function clickFilterClear()" & vbCrLf
	s = s & t2 & "{" & vbCrLf
	For i = 1 to 5
		If FilterCategory(i) > "" Then
			s = s & t2 & "for (i = 0; i < frm.new_" & FilterCategory(i) & ".length; i++)" & vbCrLf & _
				t3 & "{" & vbCrLf & _
				t3 & "frm.new_" & FilterCategory(i) & "[i].checked = false;" & vbCrLf & _
				t3 & "}" & vbCrLf
		End If
	Next
	s = s & t2 & "clickFilterGo();" & vbCrLf
	s = s & t2 & "}" & vbCrLf & vbCrLf

	s = s & t1 & "function preferencesRead()" & vbCrLf
	s = s & t2 & "{" & vbCrLf
	For i = 1 to 5
		If FilterCategory(i) > "" Then
			s = s & t2 & "value" & FilterCategory(i) & " = '';" & vbCrLf & _
				t3 & "for (i = 0; i < frm.new_" & FilterCategory(i) & ".length; i++)" & vbCrLf & _
				t3 & "{" & vbCrLf & _
					t3 & "if (frm.new_" & FilterCategory(i) & "[i].checked)" & vbCrLf & _
						t4 & "{" & vbCrLf & _
						t4 & "value" & FilterCategory(i) & " = value" & FilterCategory(i) & " + '1';" & vbCrLf & _
						t4 & "}" & vbCrLf & _
					t3 & "else" & vbCrLf & _
						t4 & "{" & vbCrLf & _
						t4 & "value" & FilterCategory(i) & " = value" & FilterCategory(i) & " + '0';" & vbCrLf & _
						t4 & "}" & vbCrLf & _
				t3 & "}" & vbCrLf & vbCrLf & _
				
			t2 & "value" & FilterCategory(i) & " = value" & FilterCategory(i) & " + " & FiftyZeros & ";" & vbCrLf & _
			t2 & "value" & FilterCategory(i) & " = value" & FilterCategory(i) & ".substring(0,49);" & vbCrLf & vbCrLf
		End If
	Next
	s = s & t2 & "}" & vbCrLf & vbCrLf
	
	s = s & t1 & "function clickFilterGo()" & vbCrLf & _
			t2 & "{" & vbCrLf & _
			t2 & "preferencesRead();" & vbCrLf & _
			t2 & "preferencesPost(" 
				For i = 1 to 5
					If FilterCategory(i) > "" Then
						s = s & "value" & FilterCategory(i) & ", "
					End If
				Next
				If Right(s, 2) = ", " Then
					s = left(s, Len(s)-2)
				End If
			s = s & ");" & vbCrLf
	s = s & t2 & "}" & vbCrLf & vbCrLf

	FunctionsFilter = FunctionsFilter & s

  End Function
  
  Function FilterHeading	
  	FilterHeading =	"<tr valign='top'>" & _
					"  <td width='100%' style='padding-top: 5px; color: #246B44; font-size: 11pt;' colspan='2'>" & _
					"	<b>Filters</b><br>" & _
					"	  <span style='margin-left: 20px;'>" & _
					"		<a href='javascript:clickFilterGo();' class='buttonFilter'>Go</a>" & _
					"		<a href='javascript:clickFilterClear();' class='buttonFilter'>Clear</a>" & _
					"	  </span>" & _
					"	<br>" & _
					"  </td>" & _
					"</tr>"
  End Function				
    
%>

<script type="text/javascript">
  
	var httpPref = new XMLHttpRequest();
	
	function preferencesPost(value1, value2, value3, value4, value5) 
        {		
		var url = 'http://www.HarvestAmerican.info/ha_BackOffice/includes/AJAX/ha_Preferences_post.asp' +
			'?prefURL=<%=strPage%>' +
			'&data1=' + value1 +
			'&data2=' + value2 +
			'&data3=' + value3 +
			'&data4=' + value4 +
			'&Now=' + escape(Date());
			
        httpPref.open('GET', url, true);
        httpPref.onreadystatechange = preferencesPostSubmit;
        httpPref.send();
        }
		
	function preferencesPostSubmit()
		{
        if (httpPref.readyState == 4) 
			{
			location.href = '<%=strPage%>'
			}
		}

</script>