
<!--Yotpo review mail after purchase code-->


<script language="JavaScript">

	var Request = new XMLHttpRequest();

	Request.open('POST', 'https://api.yotpo.com/apps/app_key/purchases');

	Request.setRequestHeader('Content-Type', 'application/json');

	Request.onreadystatechange = function () 
		{
		if (this.readyState === 4) 
			{
			console.log('Status:', this.status);
			console.log('Headers:', this.getAllResponseHeaders());
			console.log('Body:', this.responseText);
			}
		};

var body = {
	'utoken': 'HOYoM2THmVT1VaDc4y6VTS9U6IL9DrKeXKusLwlo',
	'email': 'lucinda@HarvestAmerican.com',
	'customer_name': '<%=Request("b_fname")%>',
	'order_id': '<%=Request("order_num")%>',
	'platform': 'Shopping Cart',
	'order_date': '<%=Request("date")%>',
	'currency_iso': 'USD',
	'products': 
		{
		'11121': 
			{
			'url': 'http://example_product_url1.com',
			'name': 'product1',
			'image': 'http://example_product_image_url1.com',
			'description': 'this is the description of a product',
			'price': '100'
			},
		'11133': 
			{
			'url': 'http://example_product_url2.com',
			'name': 'product2',
			'image': 'http://example_product_image_url2.com',
			'description': 'this is another description of a different product',
			'price': '200'
			}
		}
	};

	Request.send(JSON.stringify(body));

</script>



<%
	response.write "Sent Notification To Yotpo" & vbCrLf

	'Post to Task Log
	Call TaskLogWrite("Casper10", "UploadShoppingCart.asp", "ShoppingCart_SendNotificationToYotpo.inc", "Finished")	
%>