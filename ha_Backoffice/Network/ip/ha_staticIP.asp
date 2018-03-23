<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

<html>
    <head>
        <!--#include virtual="includes/connection.asp"-->
        <!--#include virtual="ha_backoffice/includes/Config.asp"-->
        <!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->
<%
ColorTab = 8
PageHeading = "Static IPs"
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
                        <span class="newStyleHeading">Static IP Blocks</span><br>
                        <table class="boarderedtable">
                            <tr class="boarderedtable">
                                <th class="boarderedtable"><b>IP</b></th>
                                <th class="boarderedtable"><b>ISP</b></th>
                                <th class="boarderedtable"><b>Description</b></th>
                            </tr>
                            <tr class="boarderedtable">
                                <td class="boarderedtable">24.103.218.144</td>
                                <td class="boarderedtable">TimeWarner</td>
                                <td class="boarderedtable">Network Address. Subnet Mask: 255.255.255.248</td>
                            </tr>
                            <tr class="boarderedtable">
                                <td class="boarderedtable">24.103.218.145</td>
                                <td class="boarderedtable">TimeWarner</td>
                                <td class="boarderedtable">Default Gateway</td>
                            </tr>
                            <tr class="boarderedtable">
                                <td class="boarderedtable">24.103.218.146</td>
                                <td class="boarderedtable">TimeWarner</td>
                                <td class="boarderedtable">Camden1</td>
                            </tr>
                            <tr class="boarderedtable">
                                <td class="boarderedtable">24.103.218.147</td>
                                <td class="boarderedtable">TimeWarner</td>
                                <td class="boarderedtable">Camden2</td>
                            </tr>
                            <tr class="boarderedtable">
                                <td class="boarderedtable">24.103.218.148</td>
                                <td class="boarderedtable">TimeWarner</td>
                                <td class="boarderedtable">Camden03</td>
                            </tr>
                            <tr class="boarderedtable">
                                <td class="boarderedtable">24.103.218.149</td>
                                <td class="boarderedtable">TimeWarner</td>
                                <td class="boarderedtable">Camden04</td>
                            </tr>
                            <tr class="boarderedtable">
                                <td class="boarderedtable">24.103.218.150</td>
                                <td class="boarderedtable">TimeWarner</td>
                                <td class="boarderedtable">Camden05</td>
                            </tr>
                            <hr>
                            <tr class="boarderedtable">
                                <td class="boarderedtable">50.75.96.136</td>
                                <td class="boarderedtable">TimeWarner</td>
                                <td class="boarderedtable">Network Address. Network Mask: 255.255.255.248</td>
                            </tr>
                            <tr class="boarderedtable">
                                <td class="boarderedtable">50.75.96.137</td>
                                <td class="boarderedtable">TimeWarner</td>
                                <td class="boarderedtable">Default Gateway</td>
                            </tr>
                            <tr class="boarderedtable">
                                <td class="boarderedtable">50.75.96.138</td>
                                <td class="boarderedtable">TimeWarner</td>
                                <td class="boarderedtable"></td>
                            </tr>
                            <tr class="boarderedtable">
                                <td class="boarderedtable">50.75.96.139</td>
                                <td class="boarderedtable">TimeWarner</td>
                                <td class="boarderedtable"></td>
                            </tr>
                            <tr class="boarderedtable">
                                <td class="boarderedtable">50.75.96.140</td>
                                <td class="boarderedtable">TimeWarner</td>
                                <td class="boarderedtable"></td>
                            </tr>
                            <tr class="boarderedtable">
                                <td class="boarderedtable">50.75.96.141</td>
                                <td class="boarderedtable">TimeWarner</td>
                                <td class="boarderedtable"></td>
                            </tr>
                            <tr class="boarderedtable">
                                <td class="boarderedtable">50.75.96.142</td>
                                <td class="boarderedtable">TimeWarner</td>
                                <td class="boarderedtable"></td>
                            </tr>
                            <tr class="boarderedtable">
                                <td class="boarderedtable">66.155.174.</td>
                                <td class="boarderedtable">WindStream</td>
                                <td class="boarderedtable"></td>
                            </tr>
                            <tr class="boarderedtable">
                                <td class="boarderedtable">66.155.174.49</td>
                                <td class="boarderedtable">WindStream</td>
                                <td class="boarderedtable"></td>
                            </tr>
                            <tr class="boarderedtable">
                                <td class="boarderedtable">66.155.174.50</td>
                                <td class="boarderedtable">WindStream</td>
                                <td class="boarderedtable"></td>
                            </tr>
                            <tr class="boarderedtable">
                                <td class="boarderedtable">66.155.174.51</td>
                                <td class="boarderedtable">WindStream</td>
                                <td class="boarderedtable"></td>
                            </tr>
                            <tr class="boarderedtable">
                                <td class="boarderedtable">66.155.174.52</td>
                                <td class="boarderedtable">WindStream</td>
                                <td class="boarderedtable"></td>
                            </tr>
                            <tr class="boarderedtable">
                                <td class="boarderedtable">66.155.174.53</td>
                                <td class="boarderedtable">WindStream</td>
                                <td class="boarderedtable"></td>
                            </tr>
                            <tr class="boarderedtable">
                                <td class="boarderedtable">66.155.174.54</td>
                                <td class="boarderedtable">WindStream</td>
                                <td class="boarderedtable"></td>
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
