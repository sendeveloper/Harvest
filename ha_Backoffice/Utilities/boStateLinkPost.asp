<%
	Response.buffer=true
	Dim rs
	Dim SQL
%>

  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_Backoffice/Utilities/connection_Barley.asp"-->
  <!--#include virtual="ha_backoffice/includes/Config.asp"-->

<html>
<head>
    <title>Number-it State Regulations Link Post</title>

    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
  
    <link href="/ha_BackOffice/includes/datepick/jquery-ui.css" rel="stylesheet" />
  <link href="/ha_BackOffice/includes/ha_Backoffice.css" type="text/css" rel="stylesheet" />

</head>

	<%
	    set rs = server.createObject("ADODB.Recordset")

	    SQL="SELECT * FROM ni_StateRegulations " & _
	        "WHERE [Abbreviation] = '" & Session("State") & "'"
	    rs.Open SQL,objcon,1,3

	    rs("URL") = Request("Link")
	    rs.Update 					 
	    rs.Close
	%>

<script>
    window.opener.location.href = window.opener.location.href;
    window.close()
</script>
	

</html>
