<%
		  rs.PageSize = iRecordperpage
		  nRecordsPerPage = rs.PageSize
		  rs.CacheSize = rs.PageSize
		  intPageCount = rs.PageCount 
		  intRecordCount = rs.RecordCount		  

		  icurrentPage = TRIM(REQUEST("page"))
		  if icurrentPage = "" then 
			icurrentPage = 0
		  else
			icurrentPage = CInt(request("page"))
		  end if

		  If CInt(icurrentPage) > CInt(intPageCount) Then icurrentPage = intPageCount
		  If CInt(icurrentPage) <= 0 Then intPage = 1
		  
        If intRecordCount > 0 Then
          rs.AbsolutePage = icurrentPage + 1
		  recordCount = 1
		  Do While recordCount <= iRecordperpage And Not rs.EOF
		  
			'Dividing Separator
			If SeparatorRows = "" Then
				'Defaults
				SeparatorRows = 3
				SeparatorWidth = "3px"
				SeparatorColor = "#D8F0E3"
			End If

			If recordCount mod SeparatorRows = 0 Then
				detailStyle = "border-bottom: " + SeparatorWidth + " solid " +SeparatorColor + ";"
			Else
				detailStyle = "border: 0;"
			End If
			
			row=(iCurrentPage * iRecordperpage) + recordCount
%>