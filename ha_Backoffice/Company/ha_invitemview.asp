<!DOCTYPE html>

<html>
<head>
  <%
    PageHeading = "View Inventory Item"

    Dim rs, rsEmp

    Function iif(condition, result, alternative)
      If condition Then iif = result Else iif = alternative
    End Function
  %>
  <title>Harvest American Backoffice - <%=PageHeading%></title>
  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/Config.asp"-->
  <!--#include file="sql.asp"-->
  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
  <link href="http://code.jquery.com/ui/1.10.2/themes/south-street/jquery-ui.css" rel="stylesheet" />
  <link href="<%=strPathIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet" />
  <script type="text/javaScript" src="/ha_BackOffice/includes/ha_Functions.js"></script>
  <script type="text/javascript">
      function validate() {
        var message = '';
        var count = 0;
        var form = document.frm;

        if (form.RequestedBy.value === '') {
            message = 'You need to enter a person whom is making the request.';
            count++;
        }
        if (form.Reason.value === '') {
            if (count > 0) {message += '\n';}
            message += 'You need to enter a reason for the request.';
            count++;
        }
        if (form.RequestFrom.value === '') {
            if (count > 0) { message += '\n'; }
            message += 'You need to enter at least one date for the request.';
            count++;
        }

        if (count === 0) {
            document.frm.submit();
        }
        else {
            document.getElementsByID('Response').InnerHTML = message;
        }
    }
  </script>
  <style type="text/css">
    div.txtIDer
    {
      width: 35%;
      float: left;
      margin-left: 10px;
    }
    input[type="date"] { width: 40%; }
    select { width: 40%; }
    textarea
      {
        width:  40%;
        height: 60px;
        resize: none;
      }
  </style>
</head>
<body onload="SetScreen(575, 675);">
  <%
    Dim eID, eMN, eDE, eVE, eWA, eDP, eMNu, eSN, eAT, eLO, eCO, eNO, strSQL
    If Not Request("iMode") = "Add" Then
      strSQL = "SELECT * FROM ha_Inventory WHERE ID = '" & Request("id") & "'"
      Set rs = Server.CreateObject("ADODB.Recordset")
      rs.Open strSQL, connCasper10, 1, 3
      eID = rs("ID")
      eMN = rs("MachineName")
      eDE = rs("Description")
      eVE = rs("Vendor")
      eWA = rs("Warranty")
      eDP = rs("DatePurchased")
      eMNu = rs("ModelNumber")
      eSN = rs("SerialNumber")
      eAT = rs("AssignedTo")
      eLO = rs("Location")
      eCO = rs("Condition")
      eNO = rs("Notes")
      rs.close
    Else
      eID = ""
      eMN = ""
      eDE = ""
      eVE = ""
      eWA = ""
      eDP = ""
      eMNu = ""
      eSN = ""
      eAT = ""
      eLO = ""
      eCO = ""
      eNO = ""
    End If
  %>
  <form action="ha_invitempost.asp?iMode=<%=Request("iMode")%>" method="post" name="frm">
    <div style="width: 500px; margin: 0 auto">
      <h3 style="text-align: center"><% =Request("iMode") %> Inventory Item</h3><br />
      <br />
      <span style="margin: 0 5em; font-style:italic">A * indicates a required field.</span><br />
      <label id="Response" style="height: 60px"></label>
      <br />
      <% If Not Request("rMode") = "Add" Then %>
      <input type="hidden" name="ID" value="<%=eID%>" />
      <% End If %>
      <div style="width: 100%"><div class="txtIDer">Machine Name:&nbsp;*</div>
      <input type="text" name="MachineName" value="<%=eMN%>" title="If the inventory item is a computer, this is where the name goes (ex. Rice, Squash)." /></div>
      <div style="width: 100%"><div class="txtIDer" style="height: 60px">Description:&nbsp;*</div>
      <input type="text" name="Description" Requesttitle="" value="<%=eDE%>" /></div>
      <div style="width: 100%"><div class="txtIDer" style="height: 60px">Vendor:&nbsp;*</div>
      <input type="text" name="Vendor" value="<%=eMN%>" title="Name of the company the equipment was purchased from" /></div>
      <div style="width: 100%"><div class="txtIDer" style="height: 60px">Warranty:&nbsp;*</div>
      <input type="text" name="Warranty" value="<%=eMN%>" title="Length of warranty, if one exists." /></div>
      <div style="width: 100%"><div class="txtIDer" style="height: 60px">Date Purchased:&nbsp;*</div>
      <input type="date" name="DatePurchased" value="<%=iif(isNull(eDP), "", eDP)%>" title="Date that the equipment was purchased.  Check the invoice that came with it." class="datepick" /></div>
      <!--#include virtual="ha_backoffice/includes/datepick.inc"-->
      <div style="width: 100%"><div class="txtIDer" style="height: 60px">Model Number:&nbsp;*</div>
      <input type="text" name="ModelNumber" value="<%=eMNu%>" title="" /></div>
      <div style="width: 100%"><div class="txtIDer" style="height: 60px">Serial Number:&nbsp;*</div>
      <input type="text" name="SerialNumber" value="<%=eSN%>" title="" /></div>
      <div style="width: 100%"><div class="txtIDer">Assigned To:</div>
      <select name="AssignedTo">
        <option value="NULL">(Please Select)</option>
        <% 
          Set rsEmp = Server.CreateObject("ADODB.RecordSet")
          rsEmp.Open "SELECT FirstName, LastName FROM ha_EmpAccounts WHERE [ActiveStatus] = 1", connCasper10, 1, 3

          Do While Not rsEmp.EOF
            response.write "        <option value=""" & rsEmp("FirstName") & """" & iif(eRB = rsEmp("FirstName"), " selected=""selected""", "") & ">" & rsEmp("FirstName") & " " & Left(rsEmp("LastName"),1) & ".</option>"
            rsEmp.MoveNext
          Loop 
          
          rsEmp.Close
          Set rsEmp = Nothing
        %>
      </select></div>
      <div style="width: 100%"><div class="txtIDer" style="height: 60px">Location:&nbsp;*</div>
      <select name="Location">
        <option value="NULL">(Please Select)</option>
        <%
          Dim rsLoc: Set rsLoc = Server.CreateObject("ADODB.Recordset")
          rsLoc.Open "SELECT * FROM ha_Types_repl WHERE Class = 'ITInventory - Location'", 1, 3

          Do While Not rsLoc.EOF
            Response.write "        <option value=""" & rsLoc("Value") & """>" & rsLoc("Description") & "</option>"
          Loop
          rsLoc.Close
          Set rsLoc = Nothing
        %>
      </select></div>
      <div style="width: 100%"><div class="txtIDer" style="height: 60px">Condition:&nbsp;*</div>
      <select name="Condition">
        <option value="NULL">(Please Select)</option>
        <%
          Dim rsCon: Set rsCon = Server.CreateObject("ADODB.Recordset")
          rsCon.Open "SELECT * FROM ha_Types_repl WHERE Class = 'ITInventory - Condition'", 1, 3

          Do While Not rsCon.EOF
            Response.write "        <option value=""" & rsCon("Value") & """>" & rsCon("Description") & "</option>"
          Loop
          rsCon.Close
          Set rsCon = Nothing
        %>
      </select></div>
      <div style="width: 100%"><div class="txtIDer" style="height: 60px">Notes:&nbsp;*</div>
      <textarea name="Notes" rows="1" cols="1"><%=eNO %></textarea><br />
      <a href="javascript:validate();" style="margin: 0 10em" class="button">Submit</a>
    </div>
  </form>
</body>
</html>
<%
  Function DateFormat(aDate)
    ' declare variables
    Dim result, d, m, y
      
    m = Month(aDate)
    d = Day(aDate)
    y = Year(aDate)
    If m < 10 Then
      m = "0" & m
    End If

    If d < 10 Then
      d = "0" & d
    End If
    
    result = y & "-" & m & "-" & d

    DateFormat = result
  End Function
%>