<?php

$strServer = "db.Zip2Tax.com";
$strDBUsername = "z2t_link";
$strDBPassword = "H^2p6~r";
$strDatabase = "zip2tax";


$conn = mysql_connect($strServer, $strDBUsername, $strDBPassword, 0, 65536) 
    or die("Failed to connect to MySQL server on $strServer");

mysql_select_db($strDatabase, $conn)
    or die("Could not open database $strDatabase");


//Set-up query variables
$strZipCode = "94133";
$strUserName = "alanalea";
$strUserPassword = "aya!37ds";


//Execute
if($result = mysql_query( "CALL zip2tax.z2t_lookup_extended('" . $strZipCode . "','" . $strUserName . "', '" . $strUserPassword . "')" )) {	
	//Read the result
	$row = mysql_fetch_array($result, MYSQL_ASSOC);
	$response = Array(
		'success' => true,
		'rate' => $row['Sales_Tax_Rate']
	);
}
else {
	$response = Array(
		'success' => false,
		'rate' => 0
	);
}

mysql_free_result($result);

//Close the Database
mysql_close($conn);

print_r($response);

?>