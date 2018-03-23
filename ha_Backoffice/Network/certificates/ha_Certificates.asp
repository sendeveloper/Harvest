<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

<html>
    <head>
        <!--#include virtual="includes/connection.asp"-->
        <!--#include virtual="ha_backoffice/includes/Config.asp"-->
        <!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->
<%
ColorTab = 8
PageHeading = "Certificates"
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
                        <span class="newStyleHeading">SSL Certificate Authority</span>
                        <p>
                        <a href="https://www.startssl.com">StartSSL</a><br>
                        <a href="">Bootstrap Client Certificate</a><br>
                        StartSSL is a free SSL certificate authority that offers unlimited, personal, single site, SSL certificates.<br>
                        These certificates can be used for business purposes, they just don't link to business information.<br>
                        They have to have information from a specific person.<br><br>
                        After an account is created they provide you with a Bootstrap Client Certificate (Linked Above)<br>
                        This certificate is used to log into their site and to generate more certificates.<br>
                        In order to generate more certificates from the account the Bootstrap Certificate must be installed in the<br>
                        Web Browser of the PC being used.
                        </p>
                    <hr>
                    </td>
                </tr>
                <tr>
                    <td>
                        <span class="newStyleHeading">Prerequisites</span>
                        <ol>
                            <li>Credentials for one of the following email accounts for the domain to be secured.<br>
                                <ul>
                                    <li>webmaster@<i>domain</i></li>
                                    <li>hostmaster@<i>domain</i></li>
                                    <li>postmaster@<i>domain</i></li>
                                </ul>
                            </li><br>
                            <li>Credentials for the server that the domain is hosted on.</li><br>
                            <li>An account at <a href="https://www.startssl.com">Start SSL</a></li><br>
                        </ol>
                    <hr>
                    </td>
                </tr>
                <tr>
                    <td>
                    <span class="newStyleHeading">Procedure to Generate Bootstrap Client Certificate</span>
                    <ol>
                        <li>Go to <a href="http://www.startssl.com">Start SSL</a> and in the top right corner click <b>Sign-up For Free</b>.</li><br>
                        <li>Follow the steps they describe to create an account and enter the email address you want linked to the certificates.</li><br>
                        <li>After submitting the application you may be given instant access, if not you have to wait for a registration email and code to be sent to you.</li><br>
                        <li>When completely registered <a href="http://www.startssl.com">StartSSL</a> will take you through an initial setup process to generate the bootstrap certificate.</li><br>
                        <li>When prompted click <b>Install</b> to install the newly created bootstrap certificate to your browser.</li><br>
                        <li><b>IMPORTANT:</b> Make sure you <b>SAVE</b> a copy of the generated certificate. If it is lost the created account is useless.</li><br>
                        <li>The method of saving is browser dependent. Instructions can be found at: <a href="https://www.startssl.com/?app=25#4">StartSSL FAQ</a></li><br>
                    <hr>
                    </td>
                </tr>
                <tr>
                    <td>
                        <span class="newStyleHeading">Procedure to Validate a Domain Name</span><br>
                        <ol>
                            <li>Open a Web Browser with the bootstrap certificate installed in it and go to <a href="http://www.startssl.com">StartSSL</a></li><br>
                            <li>From the StartSSL Home page click on <a href="https://www.startssl.com/?app=12">Control Panel</a></li><br>
                            <li>From the <b>Control Panel</b> click on <b>Authenticate</b></li><br>
                            <li>Next click on <b>Validations Wizard</b></li><br>
                            <li>For <b>Type</b> select <b>Domain Name validation</b></li><br>
                            <li>In the box enter the <b>Top Level</b> domain of the site you want to validate and select the proper suffix.</li><br>
                            <li>On the next page select the verification email address from the ones provided.</li><br>
                            <li>When the email arrives enter the provided <b>Verification Code</b> in the box and continue.</li><br> 
                        </ol>
                    <hr>
                    </td>
                <tr>
                    <td>
                        <span class="newStyleHeading">Procedure to Generate SSL Certificate Request</span><br>
                        <ol>
                            <li>Open <b>Administrator Tools</b> and then click on <b>Internet Services Manager</b></li><br>
                            <li>In the <b>Internet Information Services (IIS) Manager</b> window, select the server.</li><br>
                            <li>Scroll to the bottom and then double-click <b>Server Certificates</b></li><br>
                            <li>From the <b>Actions</b> panel on the right, click <b>Create Certificate Request...</b></li><br>
                            <li>In the <b>Create Certificate Request</b> fill out the required information.The <b>Common Name</b> must match exactly the domain name you are trying to secure.</li><br>
                            <li>On the next screen leave the <b>Cryptographic Service Provider</b> at the default value and make sure the bit length is at least 2048 and press <b>Next</b></li><br>
                            <li>On the <b>File Name</b> window enter a unique <b>Friendly Name</b> for the certificate file, save it in a known location and then click <b>OK</b></li><br>
                        </ol>
                    <hr>
                    </td>
                </tr>
                <tr>
                    <td>
                        <span class="newStyleHeading">Procedure to Generate SSL Certificate</span><br>
                        <ol>
                            <li>Open a Web Browser with the bootstrap certificate installed in it and go to <a href="http://www.startssl.com">StartSSL</a></li><br>
                            <li>From the StartSSL Home page click on <a href="https://www.startssl.com/?app=12">Control Panel</a></li><br>
                            <li>From the <b>Control Panel</b> click on <b>Authenticate</b></li><br>
                            <li>Next click on the <b>Certificates Wizard</b></li><br>
                            <li>For <b>Certificate Target</b> select <b>Web Server SSL/TLS Certificate</b></li><br>
                        </ol>
                    <hr>
                    </td>
                </tr>
                <tr>
                    <td align="center">
                        <span class="newStyleHeading"> Procedure to Install SSL Certificate to a Server</span>
                        <ol>
                            <li>Open <b>Administrative Tools</b> and click on <b>Internet Information Services (IIS)</b></li><br>
                            <li>Click on the name of the server in the connections column on the left then double click on <b>Server Certificates</b></li><br>
                            <li>In the <b>Actions</b> column on the right, click on <b>Complete Certificate Request...</b></li><br>
                            <li>Enter the path to the response certificate from the certificate authority. It should have the .cer extension. If it doesn't select to view all files.</li><br>
                            <li>Enter a <b>Friendly Name</b> for the certificate in the format: [CreationDate]<i>Domain</i>[ExpireDate]</li><br>
                            <li>If the installation was successful the new certificate should appear in the <b>Server Certificates</b> list in the middle panel.</li><br>
                        </ol>
                    <hr>
                    </td>
                </tr>
                <tr>
                    <td>
                    <span class="newStyleHeading">Procedure to Bind SSL Certificate to a Site</span><br>
                        <ol>
                            <li>Open <b> Administrator Tools</b> and then click on <b>Internet Services Manager</b></li><br>
                            <li>In the <b>Internet Services Manager (IIS)</b> window, select the name of the server where you installed the certificate.</li><br>
                            <li>Click <b>+</b> beside <b>Sites</b> and select the site to secure with the SSL certificate.</li><br>
                            <li>In the <b>Actions</b> panel on the right, click <b>Bindings...</b></li><br>
                            <li>Click <b>Add...</b></li><br>
                            <li>In the <b>Add Site Binding</b> window:<br>
                                <ul>
                                    <li>For <b>Type</b>, select <b>https</b></li><br>
                                    <li>For <b>IP Address</b>, select <b>All Unassigned</b>, or the IP address of the site.</li><br>
                                    <li>For <b>Port</b>, type <b>443</b></li><br>
                                    <li>For <b>SSL Certificate</b>, select the SSL certificate you installed, and them click<b>OK</b></li><br>
                                </ul>
                            </li>
                            <li>Close the <b>Site Bindings</b> and <b>IIS</b> windows after installation of all certificates is complete.</li><br>
                        </ol>
                    <hr>
                    </td>
                </tr>
                <tr>
                    <td>
                        <span class="newStyleHeading">Sites with SSL Certificates</span><br>
                        <ul>
                            <li><a href="https://zip2tax.info">Zip2Tax.info</a></li><br>
                            <li><a href="https://webmail.harvestamerican.net">HarvestAmerican.net (Link Points to WebMail)</a></li><br>
                        </ul>
                    <hr>
                    </td>
                </tr>
                </tr>
                </table>
            </td> 
        </tr>
        </table>
        <!--#include virtual="ha_BackOffice/includes/BodyParts/ha_PageTail.inc"-->  
    </body>
</html>
