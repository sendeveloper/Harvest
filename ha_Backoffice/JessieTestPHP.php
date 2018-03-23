<?php
$username = "sample";
$password = "password";
$zipcode = "90210";
$url = "https://api.zip2tax.com/TaxRate-USA.xml?username=" . $username . "&password=" . $password . "&zip=" . $zipcode;
$curl = curl_init($url);
curl_setopt($curl, CURLOPT_RETURNTRANSFER, true);
$curl_response = curl_exec($curl);
/*echo $curl_response;*/
curl_close($curl);

$curl_response = str_replace("utf-16","utf-8",$curl_response); 
$doc = new DOMDocument();
$doc->loadXML($curl_response);

$addrs = $doc->getElementsByTagName( "address" );

foreach ($addrs as $addr):
        $temp = $addr->getElementsByTagName( "zipCode" );
        $zip= $temp->item(0)->nodeValue;
        
        $temp = $addr->getElementsByTagName( "taxRate" );
        $salesTax= $temp->item(0)->nodeValue;
        
        $temp = $addr->getElementsByTagName( "place" );
        $place= $temp->item(0)->nodeValue;
        
        $temp = $addr->getElementsByTagName( "county" );
        $county= $temp->item(0)->nodeValue;
        
        $temp = $addr->getElementsByTagName( "state" );
        $state= $temp->item(0)->nodeValue;
endforeach;

$notes = $doc->getElementsByTagName( "notes" );
foreach ($addrs as $addr):
        $temp = $addr->getElementsByTagName( "note" );
        $notes= $temp->item(1)->nodeValue;
endforeach;

echo "Zip Code: $zip" ;
echo "\n";
echo "Sales Tax Rate: $salesTax" ;
echo "\n";
echo "Post Office City: $place" ;
echo "\n";
echo "County: $county" ;
echo "\n";
echo "State: $state" ;
echo "\n";
if($notes == "Shipping charges are not taxable")
{
  echo "Shipping Taxable: 1" ;
}
else
{
  echo "Shipping Taxable: 0" ;
}

echo $doc->errorInfo->errorMessage[0];

?>