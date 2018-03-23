  <!--#include virtual="ha_backoffice/includes/Config.asp"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsHelp.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->
<!--#Include  virtual="ha_backoffice/includes/ha_Functions_Notes.inc"-->
<%

If Request.QueryString("a") = "Save" Then
		strSQL = "ha_ToDo_Note_Insert(" & Request("ID") & "," & Request("LinkID") & ",'" & Request("Category") & "','" & Request("Note") & "','"  & Session("ha_shortname") &  "')"
		response.write strSql
		'connCasper10.Execute strSQL
		'response.redirect "ha_EmployeePermissions_edit.asp"
End If
	
	
%>