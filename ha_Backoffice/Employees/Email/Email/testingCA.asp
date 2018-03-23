<!DOCTYPE html>
<html lang="en">



<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8">
<script src="http://ajax.googleapis.com/ajax/libs/jquery/1.9.1/jquery.min.js"></script>
<script type='text/javascript' src="http://cdnjs.cloudflare.com/ajax/libs/jquery-ajaxtransport-xdomainrequest/1.0.1/jquery.xdomainrequest.min.js"></script>
</head>



<body>

<button type="button" onclick="loadDoc()">Test Connection</button>

<script type = "text/javascript" src = "//ajax.googleapis.com/ajax/libs/jquery/2.0.0/jquery.min.js" ></script>
			  <a href="https://www.google.com">test google hyperlink</a>
<script>
function loadDoc() {
        $.getJSON('https://api.zip2tax.com/TaxRate-USA.json?username=jessie&password=jessie1&zip=90001', function(response){
            if(response.z2tLookup.errorInfo.errorCode === 0) {
                console.log(response);
            }
        });
    }
</script>

</body>
</html>

