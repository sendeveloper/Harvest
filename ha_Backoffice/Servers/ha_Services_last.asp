<html>

<!--#include virtual="includes/connection.asp"-->

<head>
    <title>Harvest American Services - Last Occurance</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <script type="text/JavaScript" src="includes/checkDate.js"></script>
    <script type="text/JavaScript" src="datepick/ts_picker.js"></script>
    <script type="text/JavaScript" src="<%=strPathIncludes%>haFunctions.js"></script>

    <link href="<%=strPathControlPanel%>includes/HarvestAmerican.css" type="text/css" rel="stylesheet">

<%
    if Request("JobExecute") > "" then
        Set conn=Server.CreateObject("ADODB.Connection")

        SELECT Case Request("Server")
        CASE "Barley1"
            conn.Open "driver=SQL Server;server=Barley1.HarvestAmerican.net;uid=davewj2o;pwd=get2it;database=msdb"
        CASE "Barley2"
            conn.Open "driver=SQL Server;server=Barley2.HarvestAmerican.net;uid=davewj2o;pwd=get2it;database=msdb"
        Case "Dallas01"
            conn.Open "driver=SQL Server;server=Dallas01.HarvestAmerican.net;uid=davewj2o;pwd=get2it;database=ha_BackOffice"
        CASE "Camden1"
            conn.Open "driver=SQL Server;server=Camden1.HarvestAmerican.net;uid=davewj2o;pwd=get2it;database=msdb"
        END SELECT

        sql = "msdb.dbo.sp_start_job " & Request("JobExecute")
        response.write sql
        conn.execute sql
        conn.close
        response.redirect "cp_Services_last.asp"
    end if

    Response.buffer=true
    dim d1
    dim d2
    dim eDate

          if isnull(Request("StartDate")) or Request("StartDate") = "" then
            d1 = date()
          else
            d1 = Request("StartDate")
          end if

          if isnull(Request("EndDate")) or Request("EndDate") = "" then
            d2 = date()
          else
            d2 = Request("EndDate")
          end if
    eDate = d2
%>

    <script type="text/javascript">


    function clickServiceLog(TaskName)
        {
        var URL = '<%=strPathControlPanel%>cp_ServiceLog.asp' +
            '?Task=' + TaskName;
        openPopUp(URL);
        }

    function clickJobHistory(Server, job_id)
        {
        var URL = '<%=strPathControlPanel%>cp_JobHistory.asp' +
            '?server=' + Server +
            '&job_id=' + job_id;
        openPopUp(URL);
        }

    </script>

</head>

<body>
<table width="100%" align="center" border="0" cellspacing="0" cellpadding="0">

  <tr> 
    <td width="80%" align="left"> 
      <h1>
        Harvest American Services - Last Occurance
      </h1>
    </td>
    <td width="20%" align="right">
      <a href="javascript:window.open('<%=strPathControlPanel%>cp_LinkTest.asp','','scrollbars=yes,fullscreen=no,resizable=no, height=500,width=550,left=200,top=100');void(0)" title="Eye On Camden NY Chat History"  class="button">Link Test</a>
      <a href="<%=strPathControlPanel%>cp_Servers.asp" class="button">Home</a>
    </td>
  </tr>

  <tr>
    <td colspan="2">
      <h3>Data from Barley.ha_Prod.ha_TaskLog</h3>
      <table width="100%" border="1" cellspacing="0" cellpadding="2">
        <tr bgcolor="#FFFF66">
          <th width="25%">Task</th>
          <th width="25%">Last Run</th>
          <th width="50%">&nbsp;</th>
        </tr>

<%
          recordCount = 0
          LineCount = 0

          set rs=server.createObject("ADODB.Recordset")
          rs.open "ha_ControlPanel_Services_viewlast", objcon, 3, 3, 4

	    if not RS.EOF then
		    do while not rs.eof
		        recordCount = recordCount + 1
		        LineCount = LineCount + 1
		        If LineCount > 2 then
			      LineCount = 0
			      strColor = "#FFFFCC"
		        Else
			      strColor = "#FFFFFF"
		        End If
%>

        <tr bgcolor=<%=strColor%>>
          <td style="border-color: #EEEEEE;">
            <%=rs("Task")%>
          </td>
          <td style="border-color: #EEEEEE;">
              <a href="javascript:clickServiceLog('<%=rs("Task")%>')" class="Inline">
            <%=rs("LastTime")%></a>
          </td>
          <td style="border-color: #EEEEEE;">
            &nbsp;
          </td>
        </tr>

<%
	            rs.MoveNext
	        loop
        end if
        rs.close
        set rs = nothing
        objcon.close
%>

        <tr bgcolor="#FFFF66">
	    <td>
            <b>Task Count:</b>&nbsp;
		<%=recordCount%>
	    </td>
          <td>
            &nbsp;
          </td>
          <td>
            &nbsp;
          </td>
        </tr>
      </table>
    </td>
  </tr>

  <!------------------------------ Barley Jobs ------------------------------>
  <tr>
    <td colspan="2">
      <br><br>
      <h3>Barley Jobs</h3>
      <table width="100%" border="1" cellspacing="0" cellpadding="2">
	    <tr bgcolor="#FFFF66">
	  	  <th width="20%">Job Name</th>
		  <th width="30%">Description</th>
	  	  <th width="10%">Action</th>
	  	  <th width="10%">Last Run</th>
	  	  <th width="30%">Last Run Message</th>
	    </tr>

<%
    objcon.Open "driver=SQL Server;server=Barley1.HarvestAmerican.net;uid=davewj2o;pwd=get2it;database=Harvest"

          recordCount = 0
          LineCount = 0

          set rs=server.createObject("ADODB.Recordset")
          rs.open "ha_ControlPanel_Jobs_list", objcon, 3, 3, 4

	    if not RS.EOF then
		    do while not rs.eof
		        recordCount = recordCount + 1
		        LineCount = LineCount + 1
		        If LineCount > 2 then
			      LineCount = 0
			      strColor = "#FFFFCC"
		        Else
			      strColor = "#FFFFFF"
		        End If
%>

		<tr bgcolor=<%=strColor%>>
		  <td style="border-color: #EEEEEE;">
                    <%=rs("Name")%>
		  </td>
		  <td style="border-color: #EEEEEE;">
		    <%=rs("Description")%></a>
		  </td>
		  <td style="border-color: #EEEEEE;">
                    <a href="<%=strPathControlPanel%>cp_Services_last.asp?JobExecute=<%=rs("Name")%>&Server=Barley1" 
                      class="buttonInline">Execute Job</a>
                    <a href="javascript:clickJobHistory('Barley1', '<%=rs("job_id")%>')" class="buttonInline">History</a>
                  </td>
		  <td style="border-color: #EEEEEE;">
		    <%=rs("last_run")%></a>
		  </td>
		  <td style="border-color: #EEEEEE;">
		    <%=rs("last_outcome_message")%></a>
		  </td>
                </tr>

<%
	            rs.MoveNext
	        loop
        end if
        rs.close
        set rs = nothing
    objcon.close
%>

        <tr bgcolor="#FFFF66">
          <td>
            <b>Task Count:</b>&nbsp;
		<%=recordCount%>
          </td>
          <td>
            &nbsp;
          </td>
          <td>
            &nbsp;
          </td>
          <td>
            &nbsp;
          </td>
          <td>
            &nbsp;
          </td>
        </tr>
      </table>
    </td>
  </tr>


  <!------------------------------ Barley2 Jobs ------------------------------>
  <tr>
    <td colspan="2">
      <br><br>
      <h3>Barley2 Jobs</h3>
      <table width="100%" border="1" cellspacing="0" cellpadding="2">
	    <tr bgcolor="#FFFF66">
	  	  <th width="20%">Job Name</th>
		  <th width="30%">Description</th>
	  	  <th width="10%">Action</th>
	  	  <th width="10%">Last Run</th>
	  	  <th width="30%">Last Run Message</th>
	    </tr>

<%
    objcon.Open "driver=SQL Server;server=Barley2.HarvestAmerican.net;uid=davewj2o;pwd=get2it;database=Harvest"

          recordCount = 0
          LineCount = 0

          set rs=server.createObject("ADODB.Recordset")
          rs.open "ha_ControlPanel_Jobs_list", objcon, 3, 3, 4

	    if not RS.EOF then
		    do while not rs.eof
		        recordCount = recordCount + 1
		        LineCount = LineCount + 1
		        If LineCount > 2 then
			      LineCount = 0
			      strColor = "#FFFFCC"
		        Else
			      strColor = "#FFFFFF"
		        End If
%>

		<tr bgcolor=<%=strColor%>>
		  <td style="border-color: #EEEEEE;">
                    <%=rs("Name")%>
		  </td>
		  <td style="border-color: #EEEEEE;">
		    <%=rs("Description")%></a>
		  </td>
		  <td style="border-color: #EEEEEE;">
                    <a href="<%=strPathControlPanel%>cp_Services_last.asp?JobExecute=<%=rs("Name")%>&Server=Barley2"
                      class="buttonInline">Execute Job</a>
                    <a href="javascript:clickJobHistory('Barley2', '<%=rs("job_id")%>')" class="buttonInline">History</a>
                  </td>
		  <td style="border-color: #EEEEEE;">
		    <%=rs("last_run")%></a>
		  </td>
		  <td style="border-color: #EEEEEE;">
		    <%=rs("last_outcome_message")%></a>
		  </td>
                </tr>

<%
	            rs.MoveNext
	        loop
        end if
        rs.close
        set rs = nothing
    objcon.close
%>

        <tr bgcolor="#FFFF66">
          <td>
            <b>Task Count:</b>&nbsp;
		<%=recordCount%>
          </td>
          <td>
            &nbsp;
          </td>
          <td>
            &nbsp;
          </td>
          <td>
            &nbsp;
          </td>
          <td>
            &nbsp;
          </td>
        </tr>
      </table>
    </td>
  </tr>

  <!------------------------------ Barley2 Jobs ------------------------------>
  <tr>
    <td colspan="2">
      <br><br>
      <h3>Dallas01 Jobs</h3>
      <table width="100%" border="1" cellspacing="0" cellpadding="2">
	    <tr bgcolor="#FFFF66">
	  	  <th width="20%">Job Name</th>
		  <th width="30%">Description</th>
	  	  <th width="10%">Action</th>
	  	  <th width="10%">Last Run</th>
	  	  <th width="30%">Last Run Message</th>
	    </tr>

<%
    objcon.Open "driver=SQL Server;server=Dallas01.HarvestAmerican.net;uid=davewj2o;pwd=get2it;database=ha_backoffice"

          recordCount = 0
          LineCount = 0

          set rs=server.createObject("ADODB.Recordset")
          rs.open "ha_ControlPanel_Jobs_list", objcon, 3, 3, 4

	    if not RS.EOF then
		    do while not rs.eof
		        recordCount = recordCount + 1
		        LineCount = LineCount + 1
		        If LineCount > 2 then
			      LineCount = 0
			      strColor = "#FFFFCC"
		        Else
			      strColor = "#FFFFFF"
		        End If
%>

		<tr bgcolor=<%=strColor%>>
		  <td style="border-color: #EEEEEE;">
                    <%=rs("Name")%>
		  </td>
		  <td style="border-color: #EEEEEE;">
		    <%=rs("Description")%></a>
		  </td>
		  <td style="border-color: #EEEEEE;">
                    <a href="<%=strPathControlPanel%>cp_Services_last.asp?JobExecute=<%=rs("Name")%>&Server=Dallas01"
                      class="buttonInline">Execute Job</a>
                    <a href="javascript:clickJobHistory('Dallas01', '<%=rs("job_id")%>')" class="buttonInline">History</a>
                  </td>
		  <td style="border-color: #EEEEEE;">
		    <%=rs("last_run")%></a>
		  </td>
		  <td style="border-color: #EEEEEE;">
		    <%=rs("last_outcome_message")%></a>
		  </td>
                </tr>

<%
	            rs.MoveNext
	        loop
        end if
        rs.close
        set rs = nothing
    objcon.close
%>

        <tr bgcolor="#FFFF66">
          <td>
            <b>Task Count:</b>&nbsp;
		<%=recordCount%>
          </td>
          <td>
            &nbsp;
          </td>
          <td>
            &nbsp;
          </td>
          <td>
            &nbsp;
          </td>
          <td>
            &nbsp;
          </td>
        </tr>
      </table>
    </td>
  </tr>

  <!------------------------------ Camden1 Jobs ------------------------------>
  <tr>
    <td colspan="2">
      <br><br>
      <h3>Camden1 Jobs</h3>
      <table width="100%" border="1" cellspacing="0" cellpadding="2">
	    <tr bgcolor="#FFFF66">
	  	  <th width="20%">Job Name</th>
		  <th width="30%">Description</th>
	  	  <th width="10%">Action</th>
	  	  <th width="10%">Last Run</th>
	  	  <th width="30%">Last Run Message</th>
	    </tr>

<%
    objcon.Open "driver=SQL Server;server=Camden1.HarvestAmerican.net;uid=davewj2o;pwd=get2it;database=SSIS_Logs"

          recordCount = 0
          LineCount = 0

          set rs=server.createObject("ADODB.Recordset")
          rs.open "ha_ControlPanel_Jobs_list", objcon, 3, 3, 4

	    if not RS.EOF then
		    do while not rs.eof
		        recordCount = recordCount + 1
		        LineCount = LineCount + 1
		        If LineCount > 2 then
			      LineCount = 0
			      strColor = "#FFFFCC"
		        Else
			      strColor = "#FFFFFF"
		        End If
%>

		<tr bgcolor=<%=strColor%>>
		  <td style="border-color: #EEEEEE;">
                    <%=rs("Name")%>
		  </td>
		  <td style="border-color: #EEEEEE;">
		    <%=rs("Description")%></a>
		  </td>
		  <td style="border-color: #EEEEEE;">
                    <a href="<%=strPathControlPanel%>cp_Services_last.asp?JobExecute=<%=rs("Name")%>&Server=Camden1"
                      class="buttonInline">Execute Job</a>
                    <a href="javascript:clickJobHistory('Camden1', '<%=rs("job_id")%>')" class="buttonInline">History</a>
                  </td>
		  <td style="border-color: #EEEEEE;">
		    <%=rs("last_run")%></a>
		  </td>
		  <td style="border-color: #EEEEEE;">
		    <%=rs("last_outcome_message")%></a>
		  </td>
                </tr>

<%
	            rs.MoveNext
	        loop
        end if
        rs.close
        set rs = nothing
    objcon.close
%>

        <tr bgcolor="#FFFF66">
          <td>
            <b>Task Count:</b>&nbsp;
		<%=recordCount%>
          </td>
          <td>
            &nbsp;
          </td>
          <td>
            &nbsp;
          </td>
          <td>
            &nbsp;
          </td>
          <td>
            &nbsp;
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td>
      <div style="margin-left: 5em; height: 4em"><!--#include virtual="ha_BackOffice/includes/BodyParts/copyright.asp"--></div>
    </td>
  </tr>
</table>
</body>
</html>
