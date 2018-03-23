<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

<html>
    <head>
        <!--#include virtual="includes/connection.asp"-->
        <!--#include virtual="ha_backoffice/includes/Config.asp"-->
        <!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->
<%
ColorTab = 8
PageHeading = "DNS Mappings"
%>
        <title>Harvest American Backoffice - <%=PageHeading%></title>
        <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
        <link href="<%=strPathIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet">
        <link rel="stylesheet" type="text/css" href="<%=strPathIncludes%>menu/dropdowntabfiles/halfmoontabs.css">
        <link rel="stylesheet" type="text/css" href="../css/section_heading.css">
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
                    <td align="center">
                        <span class="newStyleHeading">DNS Mappings for Zip2Tax.info</span><br>
                        <table class="boarderedtable">
                            <tr class="boarderedtable">
                                <th class="boarderedtable">IP Address</th>
                                <th class="boarderedtable">Site/Server</th>
                            </tr>
                            <tr class="boarderedtable">
                                <td class="boarderedtable">208.88.49.20</td>
                                <td class="boarderedtable">statistics</td>
                            </tr>

                            <tr class="boarderedtable">
                                <td class="boarderedtable">208.88.49.22</td>
                                <td class="boarderedtable">
                                    <ul>
                                        <li>zip2tax.info</li>
                                        <li>www.zip2tax.info</li>
                                </td>
                            </tr>
                            <tr class="boarderedtable">
                                <td class="boarderedtable">67.225.195.119</td>
                                <td class="boarderedtable">
                                    <ul>
                                        <li>mail</li>
                                        <li>webmail</li>
                                    </ul>
                                </td>
                        </table>
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
