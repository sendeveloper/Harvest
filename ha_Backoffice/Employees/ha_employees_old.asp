<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include virtual="includes/connection.asp"-->
<!--#include virtual="ha_BackOffice/includes/Config.asp"-->
<!--#include file="sql.asp"-->
<!--#include file="lib.asp"-->
<%
	ColorTab = 4
	PageHeading = "Employees"
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
  <link href="calendar.css" rel="stylesheet" type="text/css" />
  <script type="text/javascript" src="calendar_date_picker.js"></script>
  <title>Harvest American Backoffice - Employees</title>
  <!--<meta http-equiv="REFRESH" content="0; url==strPathEmployees%>ha_employees_add.asp" />-->
  <link href="<%=strPathIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet" />
  <link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>menu/dropdowntabfiles/halfmoontabs.css" />
  <script type="text/javascript" src="<%=strPathIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>
  <style type="text/css">
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
        
    .textbox
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
    
    .headercell
    { 
      width: 24.78%;
      float: left; 
      font-weight: bold; 
      color: White; 
      background-color: #008000;
      height: 4em;
      vertical-align: middle;
      text-align: center;
      margin: 0 auto;
      border: 1px solid white;
    }
    .regcell
    {
      width: 24.78%;
      float: left; 
      font-weight: bold;
      border: 1px solid white;
    }
  </style>
</head>
<body class="gray_desktop">
  <table cellpadding="0" cellspacing="0" class="MainBody" style="margin: 20px auto 0;width: 1200px !important;">
    <!--align="center" taken out due to being an outdated element-->
    <tr>
      <td>
        <!--#include virtual="ha_Backoffice/includes/bodyparts/ha_PageHeading.inc"-->
      </td>
    </tr>
    <tr>
      <td>
        <div id="table" style="margin: auto">
            <table style="width: 80%; margin: 0 auto" cellpadding="5" cellspacing="5">
              <tr style="background-color: #008000;">
                <th style="color: White">Last Name</th>
                <th style="color: White">First Name</th>
                <th style="color: White">Phone Number</th>
          <%=iif(Session("ha_Permissions") = "99", Response.write("<th><span class=""noprint"" style=""color: White"">Options</span></th>"),"") %>
              </tr>
          <% 
              Dim userName, editURL, count, resp: count = 1
              strsql = "select EmpID, FirstName, LastName, Phone, Mobile from ha_BackOffice.dbo.ha_EmpAccounts where ActiveStatus = 1 And Contractor = 0 ORDER BY LastName"
              Dim rs: Set rs = Server.CreateObject("ADODB.Recordset")
              rs.open strsql, connCasper10, 2, 3
              Do While Not rs.eof
                userName = rs("FirstName")

                If count Mod 3 = 0 Then
                  resp = " style=""background-color: #E4F5E4"""
                Else
                  resp = ""
                End If
            %>
              <tr<%=resp%>>
                <td><%=rs("LastName")%></td>
                <td><%=rs("FirstName") %></td>
                <td><%
            If Not rs("Phone") = "" And Not rs("Mobile") = "" Then
              Response.write rs("Phone") & "&nbsp;&nbsp;&nbsp;" & rs("Mobile")
            ElseIf Not rs("Phone") = "" Then
              Response.write rs("Phone")
            ElseIf Not rs("Mobile") = "" Then
              Response.write rs("Mobile")
            End If
            %></td>
                <td style="text-align: center">
            <%
              editURL = "<a href=""ha_employees_add.asp?EmpID=" & rs("EmpID") & """>Edit</a>"

              Response.write(iif(Session("ha_Permissions") = "99", Response.write(editURL), "")) %>
                </td>
              </tr>
            <%
              count = count + 1
              rs.MoveNext
            Loop
            %>
            </table>
            

          </div>
          <div style="clear: both; height: 3em">&nbsp;</div>
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