#!C:\Strawberry\perl\bin\perl.exe

use strict;
use warnings;
use CGI;
use DBIx::Simple;
use XML::LibXML;

#########################################
###   configuration section           ###

sub SQL_HOST { '66.119.50.228' }
sub SQL_PORT { 7843 }
sub SQL_DATABASE { 'ha_Accounting' }
sub SQL_USER { 'ShoppingCart_write' }
sub SQL_PASSWORD { 'ha_Shop42^b' }

###   end of configuration            ###
#########################################

my $db = DBIx::Simple->connect( 'dbi:ODBC:Driver={SQL Server};Server=' . SQL_HOST . ',' . SQL_PORT . ';Database=' . SQL_DATABASE,
    SQL_USER,
    SQL_PASSWORD ) or die "Can't connect to database";

my $cgi = new CGI;

# dump request to file
#open (OUT,'>>','C:/inetpub/casper10.harvestamerican.info/ha_Backoffice/Servers/ShoppingCart/Woocommerce/dump.txt') || die;
#print OUT "--------------------------------------\n";
#$cgi->save(\*OUT);
#close OUT;

my $data = $cgi->param( 'POSTDATA' );
$data =~ s/\<\?xml[^\?\>]+\?\>//;

my $respText = undef;
my $respXML = undef;

my $xml = XML::LibXML->new();

eval {
    $xml->parse_string( $data );
};

if ( $@ ) {
    $respText = $data;
}
else {
    $respXML = $data;
}

$db->query('INSERT INTO ' . SQL_DATABASE . '.dbo.ha_UploadFromShoppingCart_raw (ResponseText, ResponseXML, CreatedDate) VALUES (?, ?, getdate())', $respText, $respXML) or die $db->error;

print $cgi->header( 'text/html' );
print $cgi->start_html("Upload");
print $cgi->h1("Uploaded:<br>" );
#print $cgi->escapeHTML($data); # dump XML to answer 
print $cgi->end_html;
