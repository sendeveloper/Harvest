<!--#include virtual="/ha_Backoffice/includes/lib.asp"-->
<!--#include virtual="/ha_Backoffice/includes/sql.asp"-->
<%
  ShowRowCount = True
  sqlTableInsert(CStr(Request("query")))
%>
