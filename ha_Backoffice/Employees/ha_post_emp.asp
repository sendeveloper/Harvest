<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//ENRequesthttp://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include file="sql.asp"-->
<!--#include file="lib.asp"-->
<!--#include virtual="includes/connection.asp"-->
<!--#include virtual="ha_backoffice/includes/Config.asp"-->
<!--#include virtual="ha_backoffice/includes/adovbs.asp"-->
<!--#include virtual="ha_backoffice/includes/PhoneForm.vbs"-->
<%
	Response.Buffer = true
    Dim rs, rs2, rs3
%>
<html>
<head>
    <title>Harvest American Backoffice - Employees</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <link href="<%=strPathIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet" />
</head>
<body>
<%
  Response.write Request("Specialties")
  Set rs = server.createobject("ADODB.Recordset")

  If Session("aMode")="Add" Then
    ' declare SQL string
    Dim strSQL: strSQL = "SELECT MAX(id) AS MaxID FROM ha_EmpAccounts"
    
    'open record set
    rs.Open strSQL, conn, adOpenKeyset, adLockOptimistic
      
    ' set the primary key based on the last record added
    If rs("MaxEmpID") <> "" Then
      Session("EmpID") = rs("MaxID") + 1
      Session("primaryid") = rs("MaxID") + 1
    Else
      Session("EmpID") = 1
      Session("primaryid") = 1
    End If
    
    ' close recordset
    rs.Close
    Set rs = Nothing

    strSQL = "ha_EmpAccounts"
    Set rs = server.createobject("ADODB.Recordset")
    ' reopen recordset with a new SQL string
    rs.Open strSQL, connCasper10, 2, 3
    ' add new record
    rs.AddNew
    rs("id") = Session("primaryid")
    rs("FirstName") = Request("FirstName")
    rs("LastName") = Request("LastName")
    rs("Email") = Request("Email")
    If Not Request("DOB") = "" Then rs("DOB") = Request("DOB")
    rs("AddressLine1") = Request("AddressLine1")
    rs("AddressLine2") = Request("AddressLine2")
    rs("City") = Request("City")
    rs("State") = Request("State")
    rs("PostalCode") = Request("PostalCode")
    rs("Phone") = RemoveFormatting(Request("Phone"))
    rs("Mobile") = RemoveFormatting(Request("Mobile"))
    rs("Skype") = Request("Skype")
    rs("Specialties") = Request("Specialties")
    rs("UserName") = Request("UserName")
    rs("Password") = Request("Password")
    rs("ha_Permissions") = Request("ha_Permissions")
    rs("ha_bo_Authorization") = Request("ha_bo_Authorization")
    rs("ActiveStatus") = Request("ActiveStatus")
    rs("ProjectSpecialist") =  Request("ProjectSpecialist")
    rs("EmpStatus") = Request("EmpStatus")
    rs("VacationAllowed") = Request("VacationAllowed")
    rs("VacationUsed") = Request("VacationUsed")
    rs("PaidTimeOffAllowed") = Request("PaidTimeOffAllowed")
    rs("PaidTimeOffUsed") = Request("PaidTimeOffUsed")
    If Not Request("HiredDate") = "" Then rs("HiredDate") = Request("HiredDate")
    If Not Request("TerminationDate") = "" Then rs("TerminationDate") = Request("TerminationDate")
    rs("permitTimeClock") = Request("permitTimeClock")
    rs("permitHAorg") = Request("permitHAorg")
    rs("Authority") = Request("Authority")
    rs("CreatedBy") = Session("ha_login")
    rs("CreatedDate") = Now
    rs.Update
    ' close recordset
    rs.close
    Set rs = Nothing
  Else
    Session("EmpID") = Request("EmpID")

    ' set SQL string
    strSQL = "SELECT * FROM ha_EmpAccounts WHERE [EmpID] = " & Session("EmpID")
    
    Set rs = server.createobject("ADODB.Recordset")
    ' open recordset
    rs.Open strSQL,connCasper10, 2, 3
    ' set fields
    rs("FirstName") = Request("FirstName")
    rs("LastName") = Request("LastName")
    rs("Email") = Request("Email")
    If Not Request("DOB") = "" Then rs("DOB") = Request("DOB")
    rs("AddressLine1") = Request("AddressLine1")
    rs("AddressLine2") = Request("AddressLine2")
    rs("City") = Request("City")
    rs("State") = Request("State")
    rs("PostalCode") = Request("PostalCode")
    rs("Phone") = RemoveFormatting(Request("Phone"))
    rs("Mobile") = RemoveFormatting(Request("Mobile"))
    rs("Skype") = Request("Skype")
    rs("Specialties") = Request("Specialties")
    rs("UserName") = Request("UserName")
    rs("Password") = Request("Password")
    rs("ha_Permissions") = Request("ha_Permissions")
    rs("ha_bo_Authorization") = Request("ha_bo_Authorization")
    rs("ActiveStatus") = Request("ActiveStatus")
    rs("ProjectSpecialist") =  Request("ProjectSpecialist")
    rs("EmpStatus") = Request("EmpStatus")
    rs("VacationAllowed") = Request("VacationAllowed")
    rs("VacationUsed") = Request("VacationUsed")
    rs("PaidTimeOffAllowed") = Request("PaidTimeOffAllowed")
    rs("PaidTimeOffUsed") = Request("PaidTimeOffUsed")
    If Not Request("HiredDate") = "" Then rs("HiredDate") = Request("HiredDate")
    If Not Request("TerminationDate") = "" Then rs("TerminationDate") = Request("TerminationDate")
    rs("permitTimeClock") = Request("permitTimeClock")
    rs("permitHAorg") = Request("permitHAorg")
    rs("Authority") = Request("Authority")
    rs("EditedBy") = Session("ha_login")
    rs("EditedDate") = Now
    
    ' update recordset and close
    rs.Update
    rs.Close

    Set rs = server.createobject("ADODB.Recordset")
    ' open recordset
    rs.Open strSQL,connCasper10, 2, 3
  End If
%>
    <script type="text/javascript">
      alert('SUCCESS');
      window.opener.location.href = window.opener.location.href;
      window.close();
    </script>
</body>
</html>
