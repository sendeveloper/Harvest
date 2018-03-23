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
                        <span class="newStyleHeading">DNS Mappings for Zip2tax.com</span><br>
                        <table class="boarderedtable">
                            <tr class="boarderedtable">
                                <th class="boarderedtable">IP Address</th>
                                <th class="boarderedtable">Site/Server</th>
                            </tr>
                            <tr class="boarderedtable">
                                <td class="boarderedtable">50.62.141.51</td>
                                <td class="boarderedtable">
                                    <ul>
                                        <li>db2</li>
                                        <li>dbweb2</li>
                                    </ul>
                                </td>
                            </tr>
                            <tr class="boarderedtable">
                                <td class="boarderedtable">66.119.50.226</td>
                                <td class="boarderedtable">
                                    <ul>
                                        <li>casper06</li>
                                        <li>Zip2Tax.com (For US_WEST, ASIA_PAC and OCEANIA)</li>
                                        <li>west</li>
                                        <li>west.tables</li>
                                        <li>west.api</li>
                                        <li>west.download</li>
                                        <li>api-west</li>
                                    </ul>
                                </td>
                            </tr>
                            <tr class="boarderedtable">
                                <td class="boarderedtable">66.119.50.228</td>
                                <td class="boarderedtable">
                                    <ul>
                                        <li>billing</li>
                                        <li>ftp.billing</li>
                                        <li>images</li>
                                    </ul>
                                </td>
                            </tr>
                            <tr class="boarderedtable">
                                <td class="boarderedtable">66.119.50.229</td>
                                <td class="boarderedtable">
                                    <ul>
                                        <li>casper09</li>
                                        <li>dev</li>
                                        <li>dev.api</li>
                                        <li>dev.download</li>
                                        <li>dev.tables</li>
                                        <li>ftp.dev</li>
                                        <li>csvtester</li>
                                        <li>state-generator</li>
                                    </ul>
                                </td>
                            </tr>
                            <tr class="boarderedtable">
                                <td class="boarderedtable">66.119.55.126</td>
                                <td class="boarderedtable">nlb.west</td>
                            </tr>
                            <tr class="boarderedtable">
                                <td class="boarderedtable">67.225.195.119</td>
                                <td class="boarderedtable">
                                    <ul>
                                        <li>mail</li>
                                        <li>webmail</li>
                                    </ul>
                                </td>
                            </tr>
                            <tr class="boarderedtable">
                                <td class="boarderedtable">68.178.202.54</td>
                                <td class="boarderedtable">legacy</td>
                            </tr>
                            <tr class="boarderedtable">
                                <td class="boarderedtable">72.52.250.187</td>
                                <td class="boarderedtable">blog</td>
                            </tr>
                            <tr class="boarderedtable">
                                <td class="boarderedtable">173.193.87.210</td>
                                <td class="boarderedtable">mysql</td>
                            </tr>
                            <tr class="boarderedtable">
                                <td class="boarderedtable">208.109.189.101</td>
                                <td class="boarderedtable">
                                    <ul>
                                        <li>db</li>
                                        <li>dbweb</li>
                                        <li>dbwidget</li>
                                        <li>mssql</li>
                                    </ul>
                                </td>
                            </tr>
                            <tr class="boarderedtable">
                                <td class="boarderedtable">208.109.95.225</td>
                                <td class="boarderedtable">mysql2</td>
                            </tr>
                            <tr class="boarderedtable">
                                <td class="boarderedtable">208.88.49.18</td>
                                <td class="boarderedtable">
                                    <ul>
                                        <li>dbphilly01</li>
                                        <li>mssqldb</li>
                                        <li>mssql.api</li>
                                        <li>mysql.api</li>
                                        <li>subscription.api</li>
                                    </ul>
                                </td>
                            </tr>
                            <tr class="boarderedtable">
                                <td class="boarderedtable">208.88.49.20</td>
                                <td class="boarderedtable">
                                    <ul>
                                        <li>zip2tax.com for (US_EAST & EUROPE)</li>
                                        <li>philly03</li>
                                        <li>east</li>
                                        <li>pinpoint</li>
                                        <li>production</li>
                                        <li>ftp.production</li>
                                        <li>download</li>
                                        <li>tables</li>
                                        <li>badge</li>
                                    </ul>
                                </td>
                            </tr>
                            <tr class="boarderedtable">
                                <td class="boarderedtable">208.88.49.22</td>
                                <td class="boarderedtable">
                                    <ul>
                                        <li>staging</li>
                                        <li>ftp.staging</li>
                                        <li>staging.download</li>
                                        <li>staging.api</li>
                                        <li>staging.tables</li>
                                        <li>google</li>
                                        <li>ftp.google</li>
                                    </ul>
                                </td>
                            </tr>
                            <tr class="boarderedtable">
                                <td class="boarderedtable">208.88.49.28</td>
                                <td class="boarderedtable">
                                    <ul>
                                        <li>api</li>
                                        <li>production.api</li>
                                    </ul>
                                </td>
                            </tr>
                            <tr class="boarderedtable">
                                <td class="boarderedtable">208.88.49.29</td>
                                <td class="boarderedtable">nlb2</td>
                            </tr>
                            <tr class="boarderedtable">
                                <td class="boarderedtable">208.88.49.30</td>
                                <td class="boarderedtable">
                                    <ul>
                                        <li>nlb</li>
                                        <li>rates</li>
                                    </ul>
                                </td>
                            </tr>
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
