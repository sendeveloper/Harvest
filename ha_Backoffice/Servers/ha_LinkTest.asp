<!--#include virtual="includes/connection.asp"-->
<%
	Response.buffer=true
	Dim rs2

    dim objcon2

    set objcon2=server.CreateObject("ADODB.Connection")

    objcon2.Open "driver=SQL Server;server=barley2.harvestamerican.net;uid=davewj2o;pwd=get2it;database=z2t_WebBackoffice"


	set rs2 = server.createObject("ADODB.Recordset")

        rs2.open "dbo.z2t_Link_test",objcon2,3,3,4
%>

<HTML>
<head>
    <TITLE>Zip2Tax - Database Link Test</TITLE>

</head>

<body>
    <center>
    <TABLE width="500", border="1" cellspacing="0", cellpadding="5" bgcolor="#F5F5F5"
      bordercolor="white">

<%
	while not rs2.eof
%>
        <tr>
          <td>
            <font color="black">
              <%=rs2("TestName")%>
            </font>
          </td>
          <td>
            <font color="black">
              <%=rs2("Pass")%>
            </font>
          </td>
        </tr>

<%
	    rs2.movenext
	wend

	rs2.Close
	objcon2.close
%>
    </table>
    </center>
  <!--#include virtual="ha_BackOffice/includes/BodyParts/ha_PageTail.inc"-->
</body>
</html>



