<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <!--#include virtual="includes/connection.asp"-->
    <!--#include virtual="ha_BackOffice/includes/Config.asp"-->
    <!--#include file="sql.asp"-->
    <!--#include file="lib.asp"-->
    <%
        Dim eID, eHID, eFN, eLN, eEm, eDOB, eAL1, eAL2, eCi, eSt, ePC, ePh, eMo, eSk, eSp, eUN, ePW, eHAP, eHBOA, eAS, eES, ePS, eCo, eVA, eVU, ePTOA, ePTOU, eHD, eTD, ePTC, ePHAO, eAu
        'sqldebug = true
        
        If IsNull(request("a")) = False Then
          Response.Cookies("EmpSort")("Active") = request("a")
          Response.Cookies("EmpSort")("Employee") = request("e")
          Response.Cookies("EmpSort").Expires = DateAdd("m",1,now)
        Else
          Response.Cookies("EmpSort")("Active") = 1
          Response.Cookies("EmpSort")("Employee") = ""
          Response.Cookies("EmpSort").Expires = DateAdd("m",1,now)
        End If
		
		If Request.Cookies("EmpSort")("Active") = "" Then
			Response.Cookies("EmpSort")("Active") = 1
		End If

        'response.write(thisPage)
        Function DateFormat(aDate)
          ' declare variables
          Dim dateArray, result

          ' split the parameter into an array
          dateArray = Split(aDate,"/")

          ' format the date
          result = dateArray(2) & "-"

          If dateArray(0) < 10 Then
            result = result & "0" & dateArray(0) & "-"
          Else
            result = result & dateArray(0) & "-"
          End If

          If dateArray(1) < 10 Then
            result = result & "0" & dateArray(1)
          Else
            result = result & dateArray(1)
          End If

          DateFormat = result
        End Function
      
      If Not Request("EmpID") = "" Then
        ' declare recordset and set to populate
        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.Open "SELECT * FROM ha_BackOffice.dbo.ha_EmpAccounts WHERE EmpID = '" + Request("EmpID") + "'", connCasper10, 1, 3
        eID = rs("EmpID"): eHID = rs("HarvestID"): eFN = rs("FirstName"): eLN = rs("LastName"): eEm = rs("Email")
        If Not rs("DOB") = "" Then eDOB = DateFormat(rs("DOB"))
        eAL1 = rs("AddressLine1"): eAL2 = rs("AddressLine2"): eCi = rs("City"): eSt = rs("State"): ePC = rs("PostalCode"): ePh = rs("Phone"): eMo = rs("Mobile"): eSk = rs("Skype"): eSp = rs("Specialties"): eUN = rs("Username"): ePW = rs("Password"): eHAP = rs("ha_Permissions"): eHBOA = rs("ha_bo_Authorization"): eAS = rs("ActiveStatus"): ePS = rs("ProjectSpecialist"): eES = rs("EmpStatus")
        If Not rs("HiredDate") = "" Then eHD = DateFormat(rs("HiredDate"))
        If Not rs("TerminationDate") = "" Then eTD = DateFormat(rs("TerminationDate"))
        ePTC = rs("permitTimeClock"): ePHAO = rs("permitHAorg"): eAu = rs("Authority"): Session("aMode") = "Edit"
        eVA = rs("VacationAllowed")
        eVU = rs("VacationUsed")
        ePTOA = rs("PaidTimeOffAllowed")
        ePTOU = rs("PaidTimeOffUsed")
        rs.close
        Set rs = Nothing
        
        
      Else
        eID = "": eHID = "": eFN = "": eLN = "": eEm = "": eDOB = "": eAL1 = "": eAL2 = "": eCi = "": eSt = "": ePC = "": ePh = "": eMo = "": eSk = "": eSp = "": eUN = "": ePW = "": eHAP = 0: eHBOA = "": eAS = "": ePS = "": eES = "": eHD = "": eTD = "": ePTC = "": ePHAO = "": eAu = "": Session("aMode") = "Add"
      End If 
      ColorTab = 4
	  PageHeading = Session("aMode") & " Employee"
    %>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <link href="<%=strPathIncludes %>datepick/jquery-ui.css" rel="stylesheet" />
    <script type="text/javascript">

        window.onload = init;

        /*  Function Name:      SortUsers
        *  Function Purpose:   To dynamically change the select list's contents based on radio button inputs
        *  Parameters:
        *       radioButton - the radio button that was selected
        *  Return:
        *  Change Log:
        *       10/8/2012: Created by H. Keller
        */
        function SortUsers() {
            try {
                var rdo1 = document.getElementById("rdoActiveSort1");
                var rdo2 = document.getElementById("rdoActiveSort2");
                var rdo3 = document.getElementById("rdoActiveSort3");
                var sel = document.getElementById("selEmpStatusSort");
                var active;
                var inactive;
                var emp;

                if (rdo1.checked) {
                    active = 1;
                }
                else if (rdo2.checked) {
                    active = 0;
                }
                else {
                    active = 2;
                }

                if (sel.value != '') {
                    emp = sel.value;
                }
                else {
                    emp = '';
                }
                
                var page = "<%=thisPage%>";
                
                page = "/ha_BackOffice/Employees/ha_employees_add.asp?a=" + active + "&e=" + emp;

                window.location = page;
            }
            catch (ex) {
                alert('An error occurred on Source: ha_Employees_Add.asp Function: SortUsers(). Error:' + ex, 0);
            }
        }

        function init() {
            var rdo1 = document.getElementById("rdoActiveSort1");
            var rdo2 = document.getElementById("rdoActiveSort2");
            var rdo3 = document.getElementById("rdoActiveSort3");
            var sel = document.getElementById("selEmpStatusSort");

            rdo1.onclick = SortUsers;
            rdo2.onclick = SortUsers;
            rdo3.onclick = SortUsers;
            sel.onchange = SortUsers;
        }
    </script>
    <title>Harvest American Backoffice - Employees</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <link href="<%=strPathIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet" />
    <link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>menu/dropdowntabfiles/halfmoontabs.css" />
    <script type="text/javascript" src="<%=strPathIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>
    <style type="text/css">
      .MainBody
	    {
	      width :            1200px;
	      border :           1px solid Black;
	      background-color : #FBFBFB;
	      margin-top :       20px;
        margin-left :      auto;
        margin-right :     auto;
	    }
	    
      .button
      {
        font-weight: bold;
        font-size: 1em;
        font-family: Verdana, Arial, Helvetica, sans-serif;
        color: #336600;
        padding: .25em .5em;
        background-color: #CDE2A1;
        border-top: .125em solid #C0C0C0;
        border-right: .125em solid black;
        border-bottom: .125em solid black;
        border-left: .125em solid #C0C0C0;
        text-align: center;
        width: 7.5em;
        text-decoration: none;
        margin-bottom: 2em;
      }

      .button:hover
      {
        font-weight: bold;
        font-size: 1em;
        font-family: Verdana, Arial, Helvetica, sans-serif;
        color: white;
        background-color: #336600;
        border-color: black #C0C0C0 #C0C0C0 black;
      }

      .textbox, .datepick
      {
        display: inline-block;
        height: 1.5em;
        width: 13.3em;
        margin-bottom: .5em;
      }

      .checkbox
      {
        margin-bottom: .5em;
      }

      select
      {
        display: inline-block;
        height: 1.5em;
        margin-bottom: .5em;
      }
    </style>
</head>
<body class="gray_desktop">
  <table cellpadding="0" cellspacing="0" class="MainBody" style="margin: 1em auto 0;">
    <tr>
      <td>
        <!--#include virtual="ha_Backoffice/includes/bodyparts/ha_PageHeading.inc"-->
      </td>
    </tr>
    <tr>
      <td>
        <div id="table" style="margin: auto">
          <form name="frmAddEmp" action="ha_post_emp.asp" method="post" target="foo" onsubmit="window.open('','foo','width=450,height=300,status=yes,resizable=yes,scrollbars=yes')">
              <table style="width: 80%; margin: 0 auto">
                <tr>
                  <td>
                    <br />
                    <div style="width: 80%; color: black; margin: 0 6em">
                      <fieldset>
                        <legend>Field Sort</legend>
                        <div style="text-align: left; width: 40%; margin: 0 auto; float: left"> 
                          <label for="rdoListActive">Active/Inactive</label><br />
                          <input type="radio" id="rdoActiveSort1" name="ActiveSort" value="Active" <% 
                          If IsNull(Request.Cookies("EmpSort")("Active")) = False Then
                            Response.write iif(Request.Cookies("EmpSort")("Active") = 1, "checked='""checked""'","")
                          End If
                          %> />
                          &nbsp;Active&nbsp;
                          <input type="radio" id="rdoActiveSort2" name="ActiveSort" value="Inactive" <% 
                          If Not IsNull(Request.Cookies("EmpSort")("Active")) Then
                            Response.write iif(Request.Cookies("EmpSort")("Active") = 0, "checked='""checked""'","") 
                          End If  
                          %> />
                          &nbsp;Inactive&nbsp;
                          <input type="radio" id="rdoActiveSort3" name="ActiveSort" value="Both" <%
                          If Not IsNull(Request.Cookies("EmpSort")("Active")) Then 
                            Response.write iif(Request.Cookies("EmpSort")("Active") = 2, "checked='""checked""'","")
                          End If
                          %> />
                          &nbsp;Both
                        </div>
                        <div style="text-align: left; width: 40%; margin: 0 auto; float: left">
                          <label for="rdoListWorkType">Worker Type</label><br />
                          <select id="selEmpStatusSort" name="EmpStatusSort" style="width: 13.75em" onchange="window.location='<%request.servervariables("URL")%>?<%=iif(Len(Request("EmpID")) > 0, "EmpID=" & Request("EmpID") & "&aMode=Edit&","" )%>e=' + this.value">
                            <option value="">(Please Select)</option>
                            <option value="Part Time"<%=iif(Request.Cookies("EmpSort")("EmpStatus") = "Part Time" Or Request("e") = "Part Time", " selected=""selected""", "")%>>Part Time</option>
                            <option value="Full Time"<%=iif(Request.Cookies("EmpSort")("EmpStatus") = "Full Time" Or Request("e") = "Full Time", " selected=""selected""", "")%>>Full Time</option>
                            <option value="Contractor"<%=iif(Request.Cookies("EmpSort")("EmpStatus") = "Contractor" Or Request("e") = "Contractor", " selected=""selected""", "")%>>Contractor</option>
                            <option value="Intern"<%=iif(Request.Cookies("EmpSort")("EmpStatus") = "Intern" Or Request("e") = "Intern", " selected=""selected""", "")%>>Intern</option>
                      </select>
                        </div>
                        <div style="width: 15%;float: right">
                          <b>Choose Person:</b><br />
                          <%'Response.write "EXEC ha_Employee_Select " & Request.Cookies("EmpSort")("Active") & " , '" & Request.Cookies("EmpSort")("Employee") & "'"%>
                          <select id="selUserMenu" onchange="window.location='<%request.servervariables("URL")%>?EmpID=' + this.value + '&aMode=Edit'">
                            <option value="">(Select A User)</option>
                          <% 
                          Dim userName, rs: Set rs = Server.CreateObject("ADODB.Recordset")
                          rs.Open "EXEC ha_Employee_Select " & Request.Cookies("EmpSort")("Active") & " , '" & Request.Cookies("EmpSort")("Employee") & "'",connCasper10,1,3
                          Do While Not rs.eof
                            userName = rs("FirstName") & " " & left(rs("LastName"), 1) & "."
                          
                            Response.write "                            <option value=""" & rs("EmpID") & """" 
                            If Request("EmpID") <> "" Then
                              If CInt(rs("EmpID")) = CInt(Request("EmpID")) Then Response.write " selected=""selected""" 
                            End If
                            Response.write ">" & userName & "</option>"& chr(13)
                          
                            rs.MoveNext
                          Loop
                          %>
                          </select>
                        </div>
                      </fieldset>
                    </div>
                    <label id="lblDataValidation" style="width: 100%; text-align: center; vertical-align: middle; color: red;">&nbsp;</label>
                    <br />
                  </td>
                </tr>
                <tr><td>&nbsp;</td></tr>
                <tr>
                  <td>
                    <div id="col2" style="width: 40%; height: 20em; float: left; margin-left: 10em; margin-left: 3em; text-align: right; color: black">
                    <fieldset>
                    <legend>Personal Information</legend>
                    <input id="hidEmpID" type="hidden" name="EmpID" value="<%=iif(isNull(eID), "", eID)%>" />
                    <input id="hidHarvestID" type="hidden" name="HarvestID" value="<%=iif(isNull(eHID), "", eHID)%>" />
                    <label for="txtFirstName">First Name:&nbsp;</label>
                    <input id="txtFirstName" type="text" name="FirstName" class="textbox" maxlength="50" value="<%=iif(isNull(eFN), "", eFN)%>" /><br />
                    <label for="txtLastName">Last Name:&nbsp;</label>
                    <input id="txtLastName" type="text" class="textbox" name="LastName" maxlength="50" value="<%=iif(isNull(eLN), "", eLN)%>" /><br />
                    <label for="txtEmail">Email:&nbsp;</label>
                    <input id="txtEmail" type="text" class="textbox" name="Email" maxlength="50" value="<%=iif(isNull(eEm), "", eEm)%>" /><br />
                    <label for="txtDOB">Date of Birth:</label>
                    <input type="date" name="DOB" id="DOB" class="datepick" style="margin-bottom: .5em; width: 12em !important" value="<%=iif(isNull(eDOB), "", eDOB) %>" maxlength="50" /><br />
                    <label for="txtAddress">Address:&nbsp;</label>
                    <input id="txtAddress" type="text" class="textbox" name="AddressLine1" maxlength="50" value="<%=iif(isNull(eAL1), "", eAL1)%>" /><br />
                    <label for="txtAddress2">Address&nbsp;2:&nbsp;</label>
                    <input id="txtAddress2" type="text" class="textbox" name="AddressLine2" maxlength="50" value="<%=iif(isNull(eAL2), "", eAL2)%>" /><br />
                    <label for="txtCity">City:&nbsp;</label>
                    <input id="txtCity" type="text" class="textbox" name="City" maxlength="50" value="<%=iif(isNull(eCi), "", eCi)%>" /><br />
                    <label for="txtState">State:&nbsp;</label>
                    <input id="txtState" type="text" class="textbox" name="State" maxlength="2" value="<%=iif(isNull(eSt), "", eSt)%>" /><br />
                    <label for="txtZIP">ZIP&nbsp;Code:&nbsp;</label>
                    <input id="txtZIP" type="text" class="textbox" name="PostalCode" maxlength="10" value="<%=iif(isNull(ePC), "", ePC)%>" /><br />
                    <label for="txtHomePhone">Home&nbsp;Phone:&nbsp;</label>
                    <input id="txtHomePhone" type="text" class="textbox" name="Phone" maxlength="14" value="<%=iif(isNull(ePh), "", ePh)%>" /><br />
                    <label for="txtMobilePhone">Mobile&nbsp;Phone:&nbsp;</label>
                    <input id="txtMobilePhone" type="text" class="textbox" name="Mobile" maxlength="14" value="<%=iif(isNull(eMo), "", eMo)%>" /><br />
                    <label for="txtSkype">Skype:&nbsp;</label>
                    <input id="txtSkype" type="text" class="textbox" name="skype" maxlength="50" value="<%=iif(isNull(eSk), "", eSk)%>" /><br />
                  </fieldset>
                </div>
                <div id="col4" style="width: 40%; float: left; margin-left: 5em; margin-bottom: 1.5em; text-align: right; color: black">
                  <fieldset>
                    <legend>Company Information</legend>
                    <label for="txtSpecialties" style="width: 45%; float: left; vertical-align: middle;">Extra Specialties:&nbsp;</label><br />
                    <input id="txtSpecialties" type="text" class="textbox" name="Specialties" maxlength="100" value="<%=iif(isNull(reSp), "", eSp)%>" title="Talents and traits possessed by the user." /><br />
                    <label for="txtUsername" style="width: 45%; float: left; vertical-align: middle; margin-top: .5em">Username:&nbsp;</label>
                    <input id="txtUsername" type="text" class="textbox" name="UserName" maxlength="30" value="<%=iif(isNull(eUN), "", eUN)%>" /><br />
                    <label for="txtPassword" style="width: 45%; float: left; vertical-align: middle; margin-top: .5em">Password:&nbsp;</label>
                    <input id="txtPassword" type="text" class="textbox" maxlength="30" name="Password" value="<%=iif(isNull(ePW), "", ePW)%>" /><br />
                    <label for="txtUserType" style="width: 45%; float: left; vertical-align: middle; margin-top: .5em">User&nbsp;Type:&nbsp;</label>
                      <select id="selUserType" name="ha_Permissions" title="Helps determine access level to the Harvest American backoffice." style="width: 13.75em">
                        <option value="0"<%=iif(eHAP = 0, " selected=""selected""", "")%>>None/(Select Level)</option>
                        <option value="99"<%=iif(eHAP = 99, " selected=""selected""", "")%>>Administrator</option>
                        <option value="90"<%=iif(eHAP = 90, " selected=""selected""", "")%>>Bookkeeping</option>
                        <option value="75"<%=iif(eHAP = 75, " selected=""selected""", "")%>>Regular Employee</option>
                        <option value="60"<%=iif(eHAP = 60, " selected=""selected""", "")%>>Basic</option>
                      </select>
                      <label for="rdoBOPermission1" style="width: 45%; float: left; vertical-align: middle; margin-top: .5em">Back Office Permission:&nbsp;</label><br />
                      <input id="rdoBOPermission1" type="radio" name="ha_bo_Authorization" value="" <%=iif(eHBOA ="|None|" , "checked=""checked""" , "")%> title="Helps determine access to the Number-it backoffice." />None&nbsp;
                      <input id="rdoBOPermission2" type="radio" name="ha_bo_Authorization" value="|Wholesale|" <%=iif(eHBOA = "|Wholesale|","checked=""checked""","")%> title="Helps determine access to the Number-it backoffice." />
                Wholesale&nbsp;<input id="rdoBOPermission3" type="radio" name="ha_bo_Authorization" value="|All|" <%=iif(eHBOA = "|All|", "checked=""checked""", "")%> title="Helps determine access to the Number-it backoffice." />All<br />
                      <label for="selEmpStatus" style="width: 45%; float: left; vertical-align: middle; margin-top: 1.75em">Employment Status:</label><br />
                      <select id="selEmpStatus" name="EmpStatus" style="width: 13.75em">
                        <option value="">(Please Select)</option>
                        <option value="Part Time"<%=iif(eES = "Part Time", " selected=""selected""", "")%>>Part Time</option>
                        <option value="Full Time"<%=iif(eES = "Full Time", " selected=""selected""", "")%>>Full Time</option>
                        <option value="Contractor"<%=iif(eES = "Contractor", " selected=""selected""", "")%>>Contractor</option>
                        <option value="Intern"<%=iif(eES = "Intern", " selected=""selected""", "")%>>Intern</option>
                      </select>
                      <label for="chbActive" style="width: 45%; float: left; vertical-align: middle; margin-top: .5em">Active?:&nbsp;</label><br />
                      <div style="width: 20%; margin-left: 3em; float: left"><input id="chbActive" type="radio" name="ActiveStatus" title="Is this user actively employed or contracted?" value="1" <%=iif(eAS = "True", "checked=""checked""", "")%> />&nbsp;Yes&nbsp;</div><div style="width: 20%; float: right"><input id="chbActive2" type="radio" name="ActiveStatus" title="Is this user actively employed or contracted?" value="0" <%=iif(eAS = "False", "checked=""checked""", "")%> />&nbsp;No</div>
                      <label for="chbProjectSpec" style="width: 45%; float: left; vertical-align: middle; margin-top: .5em">Project&nbsp;Specialist:&nbsp;</label><br />
                      <div style="width: 20%; margin-left: 3em; float : left"><input id="chbProjectSpec" type="radio" name="ProjectSpecialist" title="Does this user have permission to have their own project page on the HA employee page?" value="1" <%=iif(ePS = "True", "checked=""checked""", "") %> />&nbsp;Yes&nbsp;</div><div style="width: 20%; float: right"><input id="chbProjectSpec2" type="radio" name="ProjectSpecialist" title="Does this user have permission to have their own project page on the HA employee page?" value="0" <%=iif(ePS = "False", "checked=""checked""", "") %> />&nbsp;No</div>
                      <!--<label for="chbContractor" style="width: 45%; float: left; vertical-align: middle; margin-top: .5em">Employment&nbsp;Status:&nbsp;</label>
                      <div style="width: 25%; float: left"><input id="chbContractor" type="radio" name="Contractor" title="Is this user a contractor or a regular employee?" value="1" <%=iif(eCo = "True", "checked=""checked""", "") %> />&nbsp;Contractor</div><div style="width: 22%; float: right"><input id="chbContractor2" type="radio" name="Contractor" title="Is this user a contractor or a regular employee?" value="0" <%=iif(eCo = "False", "checked=""checked""", "") %> />&nbsp;Employee</div><br /><br />-->
                      <label for="txtVacation" style="width: 45%; float: left; vertical-align: middle; margin-top: 1em">Vacation Time Allowed:&nbsp;</label><br />
                      <input id="txtVacation" name="VacationAllowed" type="text" style="width: 3em; ; margin-top: .5em; text-align: right" value="<%=iif(isNull(eVA), "", eVA)%>" />&nbsp;
                      <label for="txtVacationUsed" style="margin-top: .5em">Time Used:&nbsp;</label><input id="txtVacationUsed" name="VacationUsed" type="text" style="width: 3em; margin-top: .5em; text-align: right" value="<%=iif(isNull(eVU), "", eVU)%>" /><br />
                      <label for="txtPTO" style="width: 45%; float: left; vertical-align: middle; margin-top: 1.75em">Paid Time Off Allowed?:&nbsp;</label><br />
                      <input id="txtPTO" type="text" name="PaidTimeOffAllowed" style="width: 3em; text-align: right" value="<%=iif(isNull(ePTOA), "", ePTOA)%>" />&nbsp;
                      <label for="txtPTOUsed">Time Used:&nbsp;</label><input id="txtPTOUsed" name="PaidTimeOffUsed" type="text" style="width: 3em; text-align: right" value="<%=iif(isNull(ePTOU), "", ePTOU)%>" /><br />
                      <label for="hireddate" style="width: 45%; float: left; vertical-align: middle; margin-top: 2em">Hire Date:</label><br />
                      <input type="date" name="hireddate" id="hireddate" class="datepick" value="<%=iif(isNull(eHD), "", eHD) %>" style="width: 12em" />
                      <label for="terminationdate" style="width:45%; float: left; vertical-align: middle; margin-top: .5em">Termination Date:</label>
                      <input type="date" name="terminationdate" id="terminationdate" class="datepick" value="<%=iif(isNull(eTD), "", eTD) %>" style="width: 12em" />
                      <!--#include virtual="ha_backoffice/includes/datepick.inc"--><br />
                      <label for="chbTimeClock" style="width: 45%; float: left; vertical-align: middle; margin-top: .5em">Time&nbsp;Clock&nbsp;Enabled:&nbsp;</label>
                      <div style="width: 20%; margin-left: 3em; float: left"><input id="chbTimeClock" type="radio" name="permitTimeClock" title="Grant the user the ability to record their hours for payment on the HA employee page" value="1" <%=iif(ePTC = "True", "checked=""checked""", "")%> />&nbsp;Yes&nbsp;</div><div style="width: 20%; float: right"><input id="chbTimeClock2" type="radio" name="permitTimeClock" title="Grant the user the ability to record their hours for payment on the HA employee page" value="0" <%=iif(ePTC = "False", "checked=""checked""", "")%> />&nbsp;No</div>
                      <!--<label for="chbHAOrg">HA&nbsp;Organization?:&nbsp;</label>
                      <input id="chbHAOrg" type="radio" name="permitHAorg" value="1" <%=iif(ePHAO = "True", "checked=""checked""", "")%> />&nbsp;Yes&nbsp;<input id="chbHAOrg2" type="radio" name="permitHAorg" value="0" <%=iif(ePHAO = "False", "checked=""checked""", "")%> />&nbsp;No<br />-->
                      <label for="chbAuthority" style="width: 45%; float: left; vertical-align: middle; margin-top: .5em">Authority?:&nbsp;</label>
                      <div style="width: 20%; margin-left: 3em; float: left"><input id="chbAuthority" type="radio" name="Authority" value="1" <%=iif(eAu = "True", "checked=""checked""", "")%> />&nbsp;Yes&nbsp;</div><div style="width: 20%; float: right"><input id="chbAuthority2" type="radio" name="Authority" value="0" <%=iif(eAu = "False", "checked=""checked""", "")%> />&nbsp;No</div>
                    </fieldset>
                    <br />
                    <br />
                    <input type="submit" value="Submit" class="button" style="float: right;" />
                    <input type="reset" value="Undo Changes" class="button" style="float: right; margin-right: 2em; width: 10em !important" />
                    </div>
                  </td>
                </tr>
              </table>  
              <br />
          </form>
        </div>
      </td>
    </tr>
  </table>
  <!--#include virtual="ha_BackOffice/includes/BodyParts/ha_PageTail.inc"-->
</body>
</html>

<%
Function iif(condition, consequent, alternative)
  If condition Then iif=consequent Else iif=alternative
End Function
%>
