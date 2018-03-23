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
                    <td>
                        <center><span class="newStyleHeading">DNS Mappings Index Page</span><br><hr></center>
					
                        <p>Our DNS is managed by <a href="www.dnsmadeeasy.com">DNS Made Easy</a> and the credentials are in password management.</p>
                        <br><hr>
                        <ul>
                            <li><a href="ha_dns_harvestamerican.com.asp">HarvestAmerican.com</a></li><br>
                            <li><a href="ha_dns_harvestamerican.info.asp">HarvestAmerican.info</a></li><br>
                            <li><a href="ha_dns_harvestamerican.net.asp">HarvestAmerican.net</a></li><br>
                            <li><a href="ha_dns_zip2tax.com.asp">Zip2Tax.com</a></li><br>
                            <li><a href="ha_dns_zip2tax.info.asp">Zip2Tax.info</li><br>
                            <li><a href="ha_dns_z2t.mobi">z2t.mobi</a></li><br>
                            <li><a href="ha_dns_numbermachinepro.com.asp">NumberMachinePro.com</a></li><br>
                            <li><a href="ha_dns_dxgc.org">dxgc.org</a></li><br>
                            <li><a href="ha_dns_porch.place.asp">Porch.place</a></li><br>
                            <li><a href="ha_dns_spacedman.com.asp">Spacedman.com</a></li><br>
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
