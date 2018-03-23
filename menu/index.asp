<!DOCTYPE html>

<%
  Session("redirect") = ""
  'ThisPage = "https://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL")
  'response.end
%>

<!--#include virtual="includes/connection.asp"-->

  <head>
    <meta http-equiv="content-type" content="text/html; charset=utf-8" />
    <title>Harvest American, Inc. | Home</title>
    <link href="<%=strPathMenuIncludes%>ha_menu_aj.css" type="text/css" rel="stylesheet" />
    
    <script type="text/javascript1.2" src="includes/js/functions.js"></script>
	
    <script type="text/javascript">	
	    function clickStatusChange()
			{
			var URL = '<%=strPathBase%>/menu/StaffStatusEdit.asp' +
				'?EmployeeID=<%=Session("ha_EmployeeID")%>';
			openPopUp(URL);
			}
    </script>


  </head>
  <body>

    <!--#include virtual="menu/includes/ha_header.inc"-->
    <!--#include virtual="menu/includes/ha_menu.inc"-->

    <div id="content">
	
      <!--#include virtual="menu/includes/ha_StaffStatus.asp"-->
	  
      <div id="columnB">
        <!--#include virtual="menu/includes/ha_domains.inc"-->
      </div><!-- columnB -->
	  
      <div id="columnA">
	  
        <h2>Mail Orders / Shipping</h2>
        <table width="80%" border="0" cellspacing="5" cellpadding="5" align="center">
          <tr valign="top">
            <table width="100%" border="0" cellspacing="5" cellpadding="9">
              <tr>
                <td align="center">
				
                  <a href="http://www.snscart.com/login.htm">
                    <img border="0" src="<%=strPathMenuImages%>scart.gif"></a><br />
                    <a href='http://www.snscart.com/login.htm'>Shopping Cart</a><br />
                    <!--<a href='https://www4.corecommerce.com/~salestaxdata802/admin/index.php'>CoreCommerce Cart</a> -->
                </td>
              </tr>
			</table>
		  </tr>
		  
          <tr valign="top">
            <table width="100%" border="0" cellspacing="5" cellpadding="9">
              <tr>        
                <td align="center">
                  <a href="https://ebc.cybersource.com/ebc/login/Login.do">
                    <img border="0" src="<%=strPathMenuImages%>cards_color.jpg" /></a><br />
                    <a href='https://ebc.cybersource.com/ebc/login/Login.do'>CyberSource</a>
				</td>
			  </tr>
			</table>
		  </tr>
		  
		</table>
	  </div><!-- columnA -->


      <div style="clear: both;">&nbsp;</div>
	</div>



        <!--#include virtual="menu/includes/ha_footer.inc"-->

    </body>
</html>
