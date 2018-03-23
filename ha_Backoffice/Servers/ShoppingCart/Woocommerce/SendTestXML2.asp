<%

    '// URL to which to send XML data
	URL = "/ha_backoffice/servers/ShoppingCart/woocommerce/uploadwoocommerce.pl"

    Response.Write "Sending XML data to " & URL & "<br/>"

    information = "<?xml version=""1.0"" encoding=""UTF-8""?><Orders><Order><OrderId>892</OrderId><OrderNumber>#892</OrderNumber><OrderDate>2014-10-20 13:19:41</OrderDate><OrderStatus>completed</OrderStatus><BillingFirstName>Hank</BillingFirstName><BillingLastName>Schrader</BillingLastName><BillingFullName>Hank Schrader</BillingFullName><BillingCompany></BillingCompany><BillingAddress1>4903 Cumbre Del Sur Ct NE</BillingAddress1><BillingAddress2></BillingAddress2><BillingCity>Albuquerque</BillingCity><BillingState>NM</BillingState><BillingPostCode>87111</BillingPostCode><BillingCountry>US</BillingCountry><BillingPhone>5558675309</BillingPhone><BillingEmail>schraderbrau@mailinator.com</BillingEmail><ShippingFirstName>Hank</ShippingFirstName><ShippingLastName>Schrader</ShippingLastName><ShippingFullName>Hank Schrader</ShippingFullName><ShippingCompany>DEA</ShippingCompany><ShippingAddress1>301 Dr Martin Luther King Jr Ave NE</ShippingAddress1><ShippingAddress2></ShippingAddress2><ShippingCity>Albuquerque</ShippingCity><ShippingState>NM</ShippingState><ShippingPostCode>87102</ShippingPostCode><ShippingCountry>US</ShippingCountry><ShippingMethodId>flat_rate:expedited</ShippingMethodId><ShippingMethod>Expedited</ShippingMethod><PaymentMethodId>cheque</PaymentMethodId><PaymentMethod>Cheque Payment</PaymentMethod><OrderDiscountTotal>0</OrderDiscountTotal><CartDiscountTotal>0</CartDiscountTotal><DiscountTotal>0</DiscountTotal><ShippingTotal>9.95</ShippingTotal><ShippingTaxTotal>0.597</ShippingTaxTotal><OrderTotal>105.95</OrderTotal><TaxTotal>6</TaxTotal><CompletedDate>2014-10-24 13:52:16</CompletedDate><CustomerNote>They're MINERALS, Marie.</CustomerNote><CustomerId>4</CustomerId><OrderLineItems><SKU>ninja-silhouette</SKU><ItemName>Ninja Silhouette</ItemName><Quantity>1</Quantity><Price>35</Price><LineTotal>35</LineTotal><Meta></Meta></OrderLineItems><OrderLineItems><SKU>ship-idea-green</SKU><ItemName>Ship Your Idea</ItemName><Quantity>1</Quantity><Price>20</Price><LineTotal>20</LineTotal><Meta>Color: Green</Meta></OrderLineItems></Order></Orders>"
	
    Set xmlhttp = server.Createobject("MSXML2.ServerXMLHTTP")
    xmlhttp.Open "POST", url, false
    xmlhttp.setRequestHeader "Content-Type", "text/xml" 
    xmlhttp.send information

    '// Report the response from the called page
	Response.Write "<br/><br/>"
    response.write "Response received:<hr/><span style='color:blue'>" & xmlhttp.ResponseText & "</span><hr/>"

%>
