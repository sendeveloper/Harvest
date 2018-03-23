    var isIE = 1;
    if (navigator.appName == "Netscape")
        { 
        isIE = 0;
        }


    function SetScreen(w, h)
        {
        window.resizeTo(w, h);
        CenterScreen(w, h);
        }

    function CenterScreen(w, h)
        {
        var leftprop, topprop;
 
        var leftvar = (window.screen.availWidth - w) / 2;
        var rightvar = (window.screen.availHeight - h) / 2;

        if(navigator.appName == "Microsoft Internet Explorer")
            {
            leftprop = leftvar;
            topprop = rightvar;
            }
        else 
            {
            leftprop = (leftvar - pageXOffset);
            topprop = (rightvar - pageYOffset);
            }

        window.moveTo(leftprop,topprop);
        }

    function openPopUp(URL)
        {
        window.open(URL,'',
            'scrollbars=yes,fullscreen=no,resizable=yes,height=10,width=10,left=10,top=10');
        }

    function Submit() 
		{
        document.frm.submit();
		}

	function PopupCenter(url, title, w, h) 
		{
		// Fixes dual-screen position Most browsers      Firefox
		var dualScreenLeft = window.screenLeft != undefined ? window.screenLeft : screen.left;
		var dualScreenTop = window.screenTop != undefined ? window.screenTop : screen.top;

		width = window.innerWidth ? window.innerWidth : document.documentElement.clientWidth ? document.documentElement.clientWidth : screen.width;
		height = window.innerHeight ? window.innerHeight : document.documentElement.clientHeight ? document.documentElement.clientHeight : screen.height;

		var left = ((width / 2) - (w / 2)) + dualScreenLeft;
		var top = ((height / 2) - (h / 2)) + dualScreenTop;
		var newWindow = window.open(url, title, 'scrollbars=yes, width=' + w + ', height=' + h + ', top=' + top + ', left=' + left);

		// Puts focus on the newWindow
		if (window.focus) 
			{
			newWindow.focus();
			}
		}	
		
	function clickPopupImage(ID)
		{
		var URL = '/ha_BackOffice/Company/DocumentMaintenance/ha_Document_Show.asp' +
			'?PhotoID=' + ID;
			PopupCenter(URL,'','900','500');
		}
		