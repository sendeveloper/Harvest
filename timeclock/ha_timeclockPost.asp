<!--#include virtual="includes/connection.asp"-->
<html>
<head>
    <title>Employee Hours Post</title>

    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    
</head>
<body>
<%
    Dim rs', SQL
    Set rs = server.createObject("ADODB.Recordset")
	
    If Session("cMode")="Add" then
	  rs.Open "SELECT * FROM ha_EmpHours", connCasper10, 1, 3
	  rs.AddNew
	  rs.Update
	  rs.Close
	
	  rs.Open "SELECT Max(ID) ID FROM ha_emphours", connCasper10, 2, 3
      Session("ID") = rs("ID")
      rs.Close
    End If
	
    SQL="SELECT * FROM ha_EmpHours WHERE [ID] = " & Session("ID")
    'Response.write SQL
    rs.Open SQL,connCasper10,1,3
    
    if Request("EmpID") <> "" and not IsNull(Request("DateWorked")) Then
        rs("EmpID") = Request("EmpID")
    end if
    
    if Request("DateWorked") <> "" and not IsNull(Request("DateWorked")) Then
        rs("DateWorked") = Request("DateWorked")
    end if

    if Request("QuantityHours") = "" or isNull(Request("QuantityHours")) then
        qh = 0
    else
        qh = Request("QuantityHours")
    end if
    rs("QuantityHours") = qh

    if Request("HoursType") <> "" and not IsNull(Request("HoursType")) Then
        rs("HoursType") = Request("HoursType")
    end if

    if Request("DatePaid") <> "" and not IsNull(Request("DatePaid")) Then
        rs("DatePaid") = Request("DatePaid")
    end if

    if Request("WorkDescription") <> "" and not IsNull(Request("WorkDescription")) Then
        rs("WorkDescription") = Request("WorkDescription")
    end if

    if Request("Mileage") <> "" and Request("Mileage") <> " " and not IsNull(Request("Mileage")) Then
        rs("Mileage") = Request("Mileage")
    else
        rs("Mileage") = NULL
    end if

    rs.Update 
    rs.Close

%>

<script type="text/javascript">
    alert("     YOUR TIME CARD HAS BEEN UPDATED");

    window.opener.location.href = window.opener.location.href;
    window.close()
</script>

<p>&nbsp;</p>
</body>	
</html>
