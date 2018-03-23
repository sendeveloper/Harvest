<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsHelp.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->
  <!--#include virtual="ha_backoffice/includes/Dates/ConvDate.inc"-->

<html>
<head>
  <%
    PageHeading = "Vacation Request"

    Dim rs, rsEmp

    Function iif(condition, result, alternative)
      If condition Then iif = result Else iif = alternative
    End Function
	
	Title = Request("rMode") & " Request Information"
  %>
  
  <title>Harvest American Backoffice - <%=PageHeading%></title>
  <!--#include file="sql.asp"-->
  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
  
  <link href="/ha_BackOffice/includes/datepick/jquery-ui.css" rel="stylesheet" />
  <link href="/ha_BackOffice/includes/ha_Backoffice.css" type="text/css" rel="stylesheet" />
  
  <link href="/ha_Backoffice/includes/ha_BackOfficePopup.css" type="text/css" rel="stylesheet">
  <script language="javascript" type="text/javascript" src="/includes/haFunctions.js"></script>

  <script type="text/javascript">

      function displayselect(value) {
          var elem = document.getElementsByName('DepartureTimeDiv');
          var elem2 = document.getElementByName('DepartureTime');

          if (value === 'Partial') {
              elem.style.display = 'block';
              elem2.style.display = 'block';
          }
          else {
              elem.style.display = 'none';
              elem2.style.display = 'none';
          }
      }
      function isApproved(value) {
          var elem = document.getElementsById('reasondiv');
          var elem2 = document.getElementsByName('AppReason');

          if (value === '') {
              elem.style.display = 'none';
              elem2.style.display = 'none';
          }
          else {
              elem.style.display = 'block';
              elem2.style.display = 'block';
          }
      } 

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
		
		function clickRemove(id) {
            var URL = 'ha_removerequest.asp?id=' + id;
            window.open(URL, '_self')
			}
		
		
  </script>
  <style type="text/css">
    div.txtIDer
    {
      width: 40%;
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

<body onLoad="SetScreen(600,550);">

  <%
    Dim eID, eRF, eRT, eRe, eRB, eRO, eIA, eD, strSQL
    If Not Request("rMode") = "Add" Then
      strSQL = "SELECT * FROM ha_EmpTimeRequest WHERE ID = '" & Request("id") & "'"
      Set rs = Server.CreateObject("ADODB.Recordset")
      rs.Open strSQL, connCasper10, 1, 3
      eID = rs("ID")
      If Not isNull(eRF) or Not eRF = "" Then eRF = DateFormat(rs("RequestFrom"))
      If Not isNull(eRT) or eRT = "NULL" Then eRT = DateFormat(rs("RequestTo")) Else eRT = DateFormat(rs("RequestFrom"))
      eRe = rs("Reason")
      eRB = rs("RequestedBy")
	  eD = rs("Duration")
      If Not isNull(eRO) or Not eRO = "" Then eRO = DateFormat(rs("RequestedOn"))
      eIA = rs("IsApproved")
      rs.close
    End If

  %>
  <form action="ha_postrequest.asp?rMode=<%=Request("rMode")%>" method="post" name="frm">
    <div style="width: 95%; margin: 0 auto">
	  <span class="popupHeading"><%=Title%></span>
	  
	  <br><br>
	  
      <span style="margin: 0 5em; font-style:italic">A * indicates a required field.</span><br />
      <label id="Response" style="height: 60px"></label>
      <br />
      <% If Not Request("rMode") = "Add" Then %>
      <input type="hidden" name="ID" value="<%=eID%>" />
      <% End If %>
      <div class="txtIDer">Requested By:</div>
      <select name="RequestedBy"<%=iif(Session("ha_Permissions") = "99", "", iif(Not Request("rMode") = "Add"," readonly=""readonly""",""))%>>
        <option value="">(Please Select)</option>
        <% 
          Set rsEmp = Server.CreateObject("ADODB.Recordset")
          rsEmp.Open "SELECT FirstName, LastName FROM ha_EmpAccounts WHERE [ActiveStatus] = 1", connCasper10, 1, 3
          Do While Not rsEmp.EOF
        %>
          <option value="<%=rsEmp("FirstName")%>"<%=iif(eRB = rsEmp("FirstName"), " selected=""selected""", "")%>><%=rsEmp("FirstName") & " " & Left(rsEmp("LastName"),1) & "."%></option>
        <%
            rsEmp.movenext
          Loop 
          
          rsEmp.Close
          Set rsEmp = Nothing
        %>
      </select><br />
      <div class="txtIDer" style="height: 60px">Reason:&nbsp;*</div>
      <textarea name="Reason" cols="1" rows="1" title=""<%If Request("rMode") = "View" And Not Session("ha_Permissions") = "99" Then Response.write " readonly=""readonly"""%>><%=eRe%></textarea><br />
      <div class="txtIDer">Date From:&nbsp;*</div>
      <input type="date" name="RequestFrom" value="<%=iif(isNull(eRF), "", eRF)%>" title="Date that your request begins.  If you are only requesting one day, you only need to fill this field in." class="datepick" /><br />
      <div class="txtIDer">Date To:</div>
      <input type="date" name="RequestTo" value="<%=iif(isNull(eRT), "", eRT)%>" title="Date that your request lasts to.  If you are only requesting one day, you can leave this blank." class="datepick" /><br />
      <div class="txtIDer">Length:</div>
      <select name="Duration" onchange="javascript:displayselect(this.value)">
        
        <% 
			'Declaring a variable to hold the opposite of whatever duration is set for the record.
			Dim eVal
			Dim eValTwo
			'Whatever the current duration (eD) set to, set the eVal variable to the opposite.
			If eD = "Partial" then 
				If ("rMode") = "View" then
					eVal = "Full"
				Elseif Request("rMode") = "Edit" then
				eVal = "Full"
				End If
				
			End If
			
			'If there's no duration chosen, set the variables to the following...
			If eD = NULL or eD = "" then
			eD="Choose a Duration"
			eVal = "Partial"
			eValTwo = "Full"
			End If
			'This worked better than adding two separate If statements for Add mode and in the case of the old 
			'requests where the duration may not have been set.
			
			'Whatever the current duration (eD) set to, set the eVal variable to the opposite.
			If eD="Full" then
				If Request("rMode") = "View" then
					eVal = "Partial"
				Elseif Request("rMode") = "Edit" then
					eVal = "Partial"
				End If
			End If	
		%>
		
		<option value=""<%= (eD)%>""><%=(eD)%></option>
        <option value="<%=(eVal)%>"> <%=(eVal)%></option>
		<option value="<%=(eValTwo)%>"> <%=(eValTwo)%></option>
      </select>
      <br />
      <div class="txtIDer" name="DepartureTimeDiv">Arrival/Departure Time:</div>
      <select name="DepartureTime">
        <option value="">(Please Select)</option>
        <option value="8:30">8:30</option>
        <option value="9:00">9:00</option>
        <option value="9:30">9:30</option>
        <option value="10:00">10:00</option>
        <option value="10:30">10:30</option>
        <option value="11:00">11:00</option>
        <option value="11:30">11:30</option>
        <option value="12:00">12:00</option>
        <option value="12:30">12:30</option>
        <option value="1:00">1:00</option>
        <option value="1:30">1:30</option>
        <option value="2:00">2:00</option>
        <option value="2:30">2:30</option>
        <option value="3:00">3:00</option>
        <option value="3:30">3:30</option>
      </select>
      <br />
      <% If Not Request("rMode") = "Add" And Session("ha_login") = "lrowlands" Then %>

      <div class="txtIDer">Date of Request:</div>
      <input type="date" name="RequestedOn" value="<%=iif(isNull(eRO), "", eRO)%>" title="Date that the request was submitted." /><br />
      <% End If %>
      <!--#include virtual="ha_backoffice/includes/datepick.inc"-->
      <% If Session("ha_Permissions") = "99" Then %>
      <div class="txtIDer">Decision:</div>
      <select name="IsApproved">
        <option value="">(Please Select)</option>
        <option value="0">Denied</option>
        <option value="1">Approved</option>
      </select>
	  

      <div style="height: 100px;">
        <div id="reasondiv" class="txtIDer">Decision Reason:</div>
        <textarea name="AppReason" cols="1" rows="1" title=""></textarea>
      </div>
      <% End If %>
	  
      <span class="popupButtons"> 
        <a href="javascript:javascript:validate();" class="bo_Button80">Submit</a>
		<% If Request ("rMode") = "Edit" OR Request("rMode") = "View" then %>
			<a href="javascript:clickRemove(<%= Request("id")%>);" class="bo_Button80">Remove</a>
		<% End If %>
        <a href="javascript:window.close();" class="bo_Button80">Cancel</a>
		
      </span>
	  
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