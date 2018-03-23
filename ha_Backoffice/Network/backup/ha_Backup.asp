<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

<html>
    <head>
        <!--#include virtual="includes/connection.asp"-->
        <!--#include virtual="ha_backoffice/includes/Config.asp"-->
        <!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->
<%
ColorTab = 8
PageHeading = "Backup"
%>
        <title>Harvest American Backoffice - <%=PageHeading%></title>
        <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
        <link href="<%=strPathIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet">
        <link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>menu/dropdowntabfiles/halfmoontabs.css">
        <link rel="stylesheet" type="text/css" href="css/section_heading.css">
        <script type="text/javascript" src="<%=strPathIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>

        <script language="javascript" type = "text/javascript">

        </script>
    </head>
    <body class="gray_desktop">
        <table align="center" cellspacing="0" cellpadding="0" class="MainBody"> 
        <tr> 
            <td>
                <!--#include virtual="ha_Backoffice/includes/bodyparts/ha_PageHeading.inc"-->
                <table width="96%" cellspacing="3", cellpadding="5" class="section">
                <tr>
                    <td>
                        <span class="newStyleHeading">What needs to be backed up?</span><br>
                        <p>
                         Main Areas to be backed up:<br>
                            <ul>
                                <li>File Servers</li>
                                <li>Database Tables</li>
                                <li>Websites</li>
                            </ul>
                            <br>
                         Backup Storage Servers:<br>
                            <ul>
                                <li>Camden06</li>
                                <li>McVille02 - Backup Tower (Z:) 5.45TB</li>
                            </ul>
                        </p>
                    <hr>
                    </td> 
                </tr>
                <tr>
                    <td>
                        <span class="newStyleHeading">Server Configuration for Backups</span>
                        Steps:
                        <ol>
                            <li>Required Software:
                                <ul>
                                    <li>Windows PowerShell</li>
                                    <li>Windows PowerShell Remote Extension (WinRm)</li>
                                </ul>
                            </li>
                            <li>Verify WinRm Service is Installed
                                <ol>Check if WinRm Service is installed using PowerShell Command: Get-Service winrm</ol>
                                <ol>If the output is not <b>Running</b> then download and install correct version for OS from Microsoft</ol>
                            </li>
                            <li>Run Command: <b>Enable-PSRemoting -force</b> this turns on PowerShell remoting.</li>
                            <li>Configure</li> 


                        </ol>
                    <hr>
                    </td>
                </tr>
                <tr>
                    <td>
                        <span class="newStyleHeading">Backing Up File Servers</span><br>
                        <p>
                            What file servers need to be backed up?<br>
                            <ul>
                                <li>Camden03 -> office-share (O:)</li>
                                <li>McVille02 -> data-drive (H:)</li>
                            </ul>
                        </p>
                    <hr>
                    <span class="subHead">How do the File Servers get backed up?</span>
                        <ol>
                            <li>Each File Server makes a local backup of the data locally.</li>
                            <li>Camden06 logs on to each and creates a zipped file which includes only files that match specified rules.</li>
                    <hr>
                    </td>
                </tr>
                </table>
            </td> 
        </tr>
        </table>
        <!--#include virtual="ha_BackOffice/includes/BodyParts/ha_PageTail.inc"-->  
    </body>
</html>
