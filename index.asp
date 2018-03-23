<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<%
  Session("redirect") = ""
%>

<!--#include virtual="includes/connection.asp"-->

  <head>
    <meta http-equiv="content-type" content="text/html; charset=utf-8" />
    <title>Harvest American, Inc. | Home</title>
    <link href="<%=strPathMenuIncludes%>ha_menu_aj.css" type="text/css" rel="stylesheet" />
    
    <!--[if lt IE 7]><script type="text/javascript" src="includes/js/iehoverfix.js"></script><![endif]-->
    <script type="text/javascript1.2" src="/includes/js/functions.js"></script>
    <script type="text/javascript">
      function clickStatusChange() {
		void;
        var URL = '/menu/StaffStatusEdit.asp' + '?EmpID=<%=Session("ha_EmpID")%>';
        openPopUp(URL);}
    </script>

  </head>
  <body>

    <!--#include virtual="menu/includes/ha_header.inc"-->
    <!--#include virtual="menu/includes/ha_menu.inc"-->

    <div id="content">
	
      <!--#include virtual="menu/includes/ha_StaffStatus.asp"-->
      <div id="columnA">
	  
        <h2>Mail Orders / Shipping</h2>
        <table width="80%" border="0" cellspacing="5" cellpadding="5" align="center">
          <tr valign="top">
            <table width="100%" border="0" cellspacing="5" cellpadding="9">
              <tr>
                <td align="center" width="95%">
				
                  <a href="http://www.snscart.com/login.htm">
                    <img border="0" src="<%=strPathMenuImages%>scart.gif"></a><br />
                    <a href='http://www.snscart.com/login.htm'>Shopping Cart</a><br />
                    <!--<a href='https://www4.corecommerce.com/~salestaxdata802/admin/index.php'>CoreCommerce Cart</a> -->
                </td>

                <td align="center" width="55%">
                  <a href="http://www.number-it.com/Home/BackOffice/boLogin.asp">
                    <img border="0" src="<%=strPathMenuImages%>logo_numberit.gif"height="40" /></a><br />
                    <a href='http://www.number-it.com'>Number-it.com</a><br />
                    <a href='http://www.number-it.com/Home/BackOffice/boLogin.asp'>Backoffice</a>
				</td>

              </tr></table></tr>
          <tr valign="top">
            <table width="100%" border="0" cellspacing="5" cellpadding="9">
              <tr>
                
                <td align="center" width="55%">
                  <a href="https://ebc.cybersource.com/ebc/login/Login.do">
                    <img border="0" src="<%=strPathMenuImages%>cards_color.jpg" /></a><br />
                    <a href='https://ebc.cybersource.com/ebc/login/Login.do'>CyberSource</a></td></tr></table></tr>
</table></tr></table></div><!-- columnA -->

      <div id="columnB">
        <!--#include virtual="menu/includes/ha_domains.inc"-->
      </div><!-- columnB -->

      <div style="clear: both;">&nbsp;</div>


      <div id="columnB2">
  <!-- -->
      </div><!-- columnB2 -->
      
    </div><!-- content -->

        <!--#include virtual="menu/includes/ha_footer.inc"-->

    </body>
</html>
