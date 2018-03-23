
<!-- Google Code for Back Office Notification Conversion Page -->
<script type="text/javascript">
/* <![CDATA[ */
var google_conversion_id = 1052565426;
var google_conversion_language = "en";
var google_conversion_format = "3";
var google_conversion_color = "ffffff";
var google_conversion_label = "aXNACLSqtQIQsr_z9QM";
var google_conversion_value = 0;
if ("total") {
  google_conversion_value = "total";
}
/* ]]> */
</script>
<script type="text/javascript" src="http://www.googleadservices.com/pagead/conversion.js">
</script>
<noscript>
<div style="display:inline;">
<img height="1" width="1" style="border-style:none;" alt="" src="http://www.googleadservices.com/pagead/conversion/1052565426/?value="total"&amp;label=aXNACLSqtQIQsr_z9QM&amp;guid=ON&amp;script=0"/>
</div>
</noscript>


<%
	response.write "Sent Notification To Google" & vbCrLf
	
	'Post to Task Log
	Call TaskLogWrite("Casper10", "UploadShoppingCart.asp", "ShoppingCart_SendNotificationToGoogle.inc", "Finished")
%>