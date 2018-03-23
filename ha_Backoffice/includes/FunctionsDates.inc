  <style type="text/css">
  
  </style>


<%
    DateCaption = "<img src=""/ha_BackOffice/includes/dates/cal.gif"" " & _
                  "width=""16"" height=""16"" border=""0"" alt=""Calendar""></a>" & _
                  "<span style=""margin-left: .5em"">[mm/dd/yyyy]</span>"
  
	Function PopupDate(id)
		PopupDate = "<a href=""javascript:show_calendar('document.frm." & ID & "', document.frm." & ID & ".value);"">" & _
		  DateCaption & "</a>"
	End Function
%>

<script language="JavaScript" src="/ha_BackOffice/includes/dates/checkDate.js" type="text/javascript"></script>
<script language="JavaScript" src="/ha_BackOffice/includes/dates/ts_picker.js" type="text/javascript"></script>    
	
<script language="JavaScript" type="text/javascript">
  

</script>