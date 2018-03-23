<script language="javascript" type = "text/javascript">

    var isIE = 1;
    if (navigator.appName == "Netscape")
        { 
        isIE = 0;
        }

function altOn(obj, t, s, b, e)
    {
    var altScreen = document.getElementById('mouseoverDiv');
    var sTop = document.body.scrollTop;
    var cursor = getPosition(e);

    altScreen.style.top = 100 + sTop;
    altScreen.style.top = (cursor.y - 20) + 'px';
    altScreen.style.left = (cursor.x + 100) + 'px';

    document.getElementById('mouseoverTitle').innerHTML = t;
    document.getElementById('mouseoverSub').innerHTML = s;
    b = b.replace(/~/g,'<br>');
    b = b.replace(/{/g,'"');
    document.getElementById('mouseoverBody').innerHTML = b;

    altScreen.style.visibility = 'visible';
    }
	
function altOnLeft(obj, t, s, b, e)
    {
    var altScreen = document.getElementById('mouseoverDiv');
    var sTop = document.body.scrollTop;
    var cursor = getPosition(e);

    //altScreen.style.top = 100 + sTop;
    altScreen.style.top = (cursor.y - 120) + 'px';
    altScreen.style.left = (cursor.x - 400) + 'px';

    document.getElementById('mouseoverTitle').innerHTML = t;
    document.getElementById('mouseoverSub').innerHTML = s;
    b = b.replace(/~/g,'<br>');
    b = b.replace(/{/g,'"');
    document.getElementById('mouseoverBody').innerHTML = b;

    altScreen.style.visibility = 'visible';
    }
	
function altOnAbove(obj, t, s, b, e)
    {
    var altScreen = document.getElementById('mouseoverDivAbove');
    var sTop = document.body.scrollTop;
    var cursor = getPosition(e);

    altScreen.style.top = (cursor.y - 140) + 'px';
    altScreen.style.left = (cursor.x - 80) + 'px';

    document.getElementById('mouseoverTitleAbove').innerHTML = t;
    document.getElementById('mouseoverSubAbove').innerHTML = s;
    b = b.replace(/~/g,'<br>');
    b = b.replace(/{/g,'"');
    document.getElementById('mouseoverBodyAbove').innerHTML = b;

    altScreen.style.visibility = 'visible';
    }
	

function altOff()
    {
    document.getElementById('mouseoverDiv').style.visibility = 'hidden';
    }

function altOffAbove()
    {
    document.getElementById('mouseoverDivAbove').style.visibility = 'hidden';
    }
	
function getPosition(e) {
    //Firefox event gets passed in, but IE needs to get it now
    if (!e) var e = window.event;
    var cursor = {x:0, y:0};
    if (e.pageX || e.pageY) {
        cursor.x = e.pageX;
        cursor.y = e.pageY;
    } 
    else {
        cursor.x = e.clientX + 
            (document.documentElement.scrollLeft || 
            document.body.scrollLeft) - 
            document.documentElement.clientLeft;
        cursor.y = e.clientY + 
            (document.documentElement.scrollTop || 
            document.body.scrollTop) - 
            document.documentElement.clientTop;
    }
    return cursor;
}


</script>

<style type="text/css">

    .mouseoverDiv
        {
		position: absolute; 
        top: 50px; 
        left: 50px; 
        width: 300px; 
        background-color: #FFFFCC;
        visibility: hidden;
        }
		
    .mouseoverDivAbove
        {
		position: absolute; 
        top: 50px; 
        left: 50px; 
        width: 200px; 
        background-color: #FFFFCC;
        visibility: hidden;
        }
		
</style>

  <div id="mouseoverDiv" class="mouseoverDiv">
    <table width='100%' border="2" cellspacing="2" cellpadding="2" bordercolor='black'>
      <tr valign='top'>
        <td>

          <table width='100%' border='0' cellspacing='0' cellpadding='0')

            <tr valign='top'>
              <td align='center' style='font-size: 9pt; padding-bottom: 5px;'>
                <b><span id='mouseoverTitle'>Topic Goes Here</span></b>
              </td>
            </tr>

            <tr valign='top'>
              <td align='center' style='font-size: 8pt; color: #C0C0C0;'>
                <b><span id='mouseoverSub'>Subtitle Goes Here</span></b>
              </td>
            </tr>
			
            <tr>
              <td align='center'>
                <hr>
              </td>
            </tr>

            <tr valign='top'>
              <td width='100%'>
                <font size='1'>
                <span id='mouseoverBody'>
                  This is where the mouseover text will go.
                </span>
                </font>
             </td>
            </tr>

          </table>

        </td>
      </tr>
    </table>
  </div>
  
<%  
Function mouseoverDivAbove()  
%>
    <div id="mouseoverDivAbove" class="mouseoverDivAbove">
    <table width='100%' border="2" cellspacing="2" cellpadding="2" bordercolor='black'>
      <tr valign='top'>
        <td>

          <table width='100%' border='0' cellspacing='0' cellpadding='0')

            <tr valign='top'>
              <td align='center' style='font-size: 9pt; padding-bottom: 5px;'>
                <b><span id='mouseoverTitleAbove'>Topic Goes Here</span></b>
              </td>
            </tr>

            <tr valign='top'>
              <td align='center' style='font-size: 8pt; color: #C0C0C0;'>
                <b><span id='mouseoverSubAbove'>Subtitle Goes Here</span></b>
              </td>
            </tr>
			
            <tr>
              <td align='center'>
                <hr>
              </td>
            </tr>

            <tr valign='top'>
              <td width='100%'>
                <font size='1'>
                <span id='mouseoverBodyAbove'>
                  This is where the mouseover text will go.
                </span>
                </font>
             </td>
            </tr>

          </table>

        </td>
      </tr>
    </table>
  </div>
<%
End Function
%>

