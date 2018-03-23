    function UserTracking(op, page, data1, data2)
        {
            uTrack.open("GET", "https://"+window.location.hostname+"/includes/UserTracking/UserTrackingPost.asp?Op=" + op + 
                "&Page=" + page +
                "&Data1=" + data1 +
                "&Data2=" + data2 +
                "&Now=" + escape(Date()), true);
            uTrack.send(null);
        }

var uTrack = new XMLHttpRequest();

