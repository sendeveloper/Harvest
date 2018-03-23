  <script type="text/javascript" src="<%=strPathMenuIncludes%>js/lib.js"></script>
  <script type="text/javascript">
    function init(event){
      get(".dropdown").forEach(function(e){
        listen(e, "mouseover", function(event){this.children.filter(function(e){return e.nodeName.toLowerCase() == "ul"}).forEach(function(e){e.style.display = e.redisplay;});}, true);
        listen(e, "mouseout", function(event){this.children.filter(function(e){return e.nodeName.toLowerCase() == "ul"}).forEach(function(e){e.style.display = "none";});}, true);
        e.children.filter(function(e){return e.nodeName.toLowerCase() == "ul"}).forEach(function(e) {e.redisplay = css(e).display; e.style.display = "none";})
        return;});
      return;}
    listen(window, "load", init);
  </script>
