function getHTTPObject() {
  var xmlhttp;
  /*@cc_on
  @if (@_jscript_version >= 5)
    try {
      xmlhttp = new ActiveXObject("Msxml2.XMLHTTP");
    } catch (e) {
      try {
        xmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
      } catch (E) {
        xmlhttp = false;
      }
    }
  @else
  xmlhttp = false;
  @end @*/
  if (!xmlhttp && typeof XMLHttpRequest != 'undefined') {
    try {
      xmlhttp = new XMLHttpRequest();
    } catch (e) {
      xmlhttp = false;
    }
  }
  return xmlhttp;
}

    function getInnerText(node) {
        if (typeof node.textContent != 'undefined') {
            return node.textContent;
            }
        else if (typeof node.innerText != 'undefined') {
            return node.innerText;
            }
        else if (typeof node.text != 'undefined') {
            return node.text;
            }
        else {
            switch (node.nodeType) {
            case 3:
            case 4:
                return node.nodeValue;
                break;
            case 1:
            case 11:
                var innerText = '';
                for (var i = 0; i < node.childNodes.length; i++) {
                    innerText += getInnerText(node.childNodes[i]);
                    }
                return innerText;
                break;
            default:
                return '';
                }
            }
        }

var http = getHTTPObject();
