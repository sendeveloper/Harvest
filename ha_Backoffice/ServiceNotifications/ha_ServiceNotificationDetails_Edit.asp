<!DOCTYPE html>
<!--#include virtual="ha_backoffice/includes/Config.asp"-->
<!--#include virtual="includes/connection.asp"-->

<%
  Session("Redirect") = ""
  ColorTab = 3
  PageHeading = "Service Notification Details"
%>

<html>
<head>
  <title>Table Distribution Status ha_EmpAccounts</title>
  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
  <link rel="stylesheet" type="text/css" href="popup.css">
</head>

<body>
  <table width="800" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr>
      <td>

        <table width="100%" border="0" cellspacing="10" cellpadding="0" align="center">

          <tr>
            <td class="popupHeading">
              Service Notification Details
            </td>
          </tr>

          <tr>
            <td>

              <table width="100%" border="0" cellspacing="2" cellpadding="2">

                <tr>
                  <td class="databaseHead" colspan="4">
                    Locations
                  </td>
                  <td></td>
                  <td class="databaseHead" colspan="6">
                    Results
                  </td>
                </tr>

                <tr>
                  <td class="subHead2" colspan="7"></td>
                  <td class="subHead2">
                    Record
                  </td>
                  <td class="subHead2">
                    Inserted
                  </td>
                  <td class="subHead2">
                    Updated
                  </td>
                </tr>

                <tr>
                  <td width="12%"></td>
                  <td class="subHead" width="10%">
                    Server
                  </td>
                  <td class="subHead" width="15%">
                    Database
                  </td>
                  <td class="subHead" width="15%">
                    Table
                  </td>
                  <td width="2%"></td>
                  <td class="subHead" width="10%">
                    Server
                  </td>
                  <td class="subHead" width="15%">
                    Database
                  </td>
                  <td class="subHead" width="7%">
                    Count
                  </td>
                  <td class="subHead" width="7%">
                    Count
                  </td>
                  <td class="subHead" width="7%">
                    Count
                  </td>
                </tr>
<%
  For i = 1 to 5
    'Set up the first column
    If i = 0 Then
      s = "Publisher"
      resultColumns = "<td></td><td></td>"
    Else
      s = "Subscriber " & cStr(i)
      resultColumns = "<td id='resultInserted" & cStr(i) & "' align='right'>- - - - -</td>" & _
                "<td id='resultUpdated" & cStr(i) & "' align='right'>- - - - -</td>"
    End If

    'Skip a line after the Publisher row
    If i = 1 Then
      Response.write "<tr><td></td><tr>"
    End If
%>

                <tr>
                  <td>
                    <span style="color: #C0C0C0;font-size: 10px;"><%=s%></span>
                  </td>
                  <td>
                    <%=LocationServer(i)%>
                  </td>
                  <td>
                    <%=LocationDatabase(i)%>
                  </td>
                  <td>
                    <%=LocationTable(i)%>
                  </td>
                  <td></td>
                  <td id="resultServer<%=i%>">
                    - - - - -
                  </td>
                  <td id="resultDatabase<%=i%>">
                    - - - - -
                  </td>
                  <td id="resultRecords<%=i%>" align="right">
                    - - - - -
                  </td>
                  <%=resultColumns%>
                </tr>

<%
  Next
%>
                <tr>
                  <td colspan="10">
                    <hr>
                  </td>
                </tr>

              </table>


              <table width="100%" border="0" cellspacing="5" cellpadding="5">
                <tr>
                  <td width="45%"></td>
                  <td align="center">
                    <a href="javascript:clearResults();" class="bo_Button100" title="Clear the Results">Clear</a>
                  </td>
                  <td align="center">
                    <a href="javascript:clickRun();" class="bo_Button100" title="Lists the status of the tables">Run Status</a>
                  </td>
                  <td align="center">
                    <a href="javascript:clickUpdate();" class="bo_Button100" title="Update all the subscriber tables to match publisher">Update</a>
                  </td>
                  <td align="center">
                    <a href="javascript:close();" class="bo_Button100" title="Closes this window">Close</a>
                  </td>
                  <td width="45%"></td>
                </tr>
              </table>

            </td>
          </tr>
        </table>

      </td>
    </tr>
  </table>

<%
  'Post the page load into activity
  'dim connPublic
  'set connPublic=server.CreateObject("ADODB.Connection")
  'connPublic.Open "driver=SQL Server;server=dbWeb.Zip2Tax.com;uid=z2t_WebUser;pwd=WebUser_z2t;database=z2t_WebPublic"

  Data2 = ""
  Data1 = request("p")

  pgeURL = Request.ServerVariables("path_info")

  If left(pgeURL,1) = "/" Then
    pgeURL = mid(pgeURL&"  ",2)
  End If

  pgeURL = trim(pgeURL)
  URL = Request.ServerVariables("HTTP_REFERER")

  SQL = "z2t_Activity_add_new('" & Session("z2t_UserName") & "', 18, " & _
          "'" & Data1 & "', " & _
          "'" & Data2 & "', " & _
          "'" & pgeURL & "', " & _
          "'" & URL & "', " & _
          "'" & Session.SessionID & "', " & _
          "'" & Request.ServerVariables("REMOTE_ADDR") & "', " & _
          "'z2t_TypesStatus.asp', " & _
          Session("CookieID") & ")"

  'connPublic.Execute(sql)
%>

</body>
</html>

<!--script type="text/javascript">
  var http = new XMLHttpRequest();
  var loopStep = 0;
  var locationCount = <%=LocationCount%>;
  var locationServer = [
    <%For i = 0 to LocationCount -1
      Response.write "'" & LocationServer(i) & "', "
    Next%>];
  var locationDatabase = [
    <%For i = 0 to LocationCount -1
      Response.write "'" & LocationDatabase(i) & "', "
    Next%>];
  var aRecMax = 0;
  var updateFlag = 0;
  var totalInserted = 0;
  var totalUpdated = 0;

    function clearResults()
        {
    loopStep = 0;
    for (i=0; i < locationCount; i++)
      {
      var eleID = 'resultServer'+i.toString();
      document.getElementById(eleID).innerHTML = '- - - - -';
      var eleID = 'resultDatabase'+i.toString();
      document.getElementById(eleID).innerHTML = '- - - - -';
      var eleID = 'resultRecords'+i.toString();
      document.getElementById(eleID).innerHTML = '- - - - -';

      if (i > 0)
        {
        var eleID = 'resultInserted'+i.toString();
        document.getElementById(eleID).innerHTML = '- - - - -';
        var eleID = 'resultUpdated'+i.toString();
        document.getElementById(eleID).innerHTML = '- - - - -';
        }
      }
    }

    function clickRun()
        {
    clearResults();
        runStatusCheck();
    }

    function clickUpdate()
        {
    clearResults();
        runStatusCheck();
    updateFlag = 1;
    }

    function formLoad()
        {
        SetScreen(850, 600);
    }

    function runStatusCheck()
        {
    var url = 'http://www.zip2tax.com/backoffice/TableDistribution_ha_empAccounts/TableDistribution_ha_empaccounts_readstatus.asp' +
      '?s=' + locationServer[loopStep] +
      '&d=' + locationDatabase[loopStep] +
      '&Now=' + escape(Date());
    //alert(url);

        http.open('GET', url, true);
        http.onreadystatechange = getResponse;
        http.send();
        }

    function getResponse()
        {
        if (http.readyState == 4)
      {
      if (http.status == 200)
        {
        var r = http.responseXML;
        var eleID = 'resultServer'+loopStep.toString();
        document.getElementById(eleID).innerHTML = r.getElementsByTagName('response_server_name')[0].firstChild.nodeValue;
        var eleID = 'resultDatabase'+loopStep.toString();
        document.getElementById(eleID).innerHTML = r.getElementsByTagName('response_database_name')[0].firstChild.nodeValue;
        var eleID = 'resultRecords'+loopStep.toString();
        document.getElementById(eleID).innerHTML = r.getElementsByTagName('response_record_count')[0].firstChild.nodeValue;

        //Set the maximum AccountID from the source response
        if (loopStep == 0)
          {
          aRecMax = r.getElementsByTagName('response_maximum_ID')[0].firstChild.nodeValue;
          }

        //Move to the next location
        if (loopStep < locationCount - 1)
          {
          loopStep = loopStep + 1;
          runStatusCheck();
          }
        else if (updateFlag == 1)
          {
          //Start the update portion
          loopStep = 1;  //We don't want to start at 0, that's the source
          runUpdate();
          }
        }
      }
        }

  function runUpdate()
        {
    var url = 'http://www.zip2tax.com/backoffice/TableDistribution_ha_EmpAccounts/TableDistribution_ha_EmpAccounts_update.asp' +
      '?s=' + locationServer[loopStep] +
      '&d=' + locationDatabase[loopStep] +
      '&Now=' + escape(Date());
    //alert(url);

        http.open('GET', url, true);
        http.onreadystatechange = getUpdateResponse;
        http.send();
        }

    function getUpdateResponse()
        {
        if (http.readyState == 4)
      {
      if (http.status == 200)
        {
        var r = http.responseXML;
        var eleID = 'resultServer'+loopStep.toString();
        document.getElementById(eleID).innerHTML = r.getElementsByTagName('response_server_name')[0].firstChild.nodeValue;
        var eleID = 'resultDatabase'+loopStep.toString();
        document.getElementById(eleID).innerHTML = r.getElementsByTagName('response_database_name')[0].firstChild.nodeValue;
        var eleID = 'resultInserted'+loopStep.toString();
        document.getElementById(eleID).innerHTML = r.getElementsByTagName('response_records_inserted')[0].firstChild.nodeValue;
        var eleID = 'resultUpdated'+loopStep.toString();
        document.getElementById(eleID).innerHTML = r.getElementsByTagName('response_records_updated')[0].firstChild.nodeValue;

        if (loopStep < locationCount - 1)
          {
          //Table finished, move to the next
          loopStep = loopStep + 1;
          runUpdate();
          }
        else
          {
          //All done
          updateFlag = 0;
          }
        }
      }
        }

   window.onLoad="formLoad();"
</script-->
