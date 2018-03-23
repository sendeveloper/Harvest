<%
  Session("Redirect") = ""
  ColorTab = 3
  PageHeading = "Service Notification Status"
%>
<!--#include virtual="ha_BackOffice/ServiceNotifications/top.inc"-->
<% 
    Dim count: count = 0
%>
<style>
 .pass {color: green;}
 .fail {color: black; background: red; font-weight: bold; text-decoration: blink;}
</style>

<script language="javascript">
  function action(place, data, replace) {
    var h = ajax("/ha_BackOffice/ServiceNotifications/ha_ServiceNotificationAjax.asp", 
      function(body) {
        get(place)[0][0].toMe(function(e){
          if (replace) {
            e.innerHTML = body;}
          else {
            e.appendChild(node("div").toMe(function(e){e.innerHTML = body;}));
          //e.appendChild(text(body));
        }})},
      function(body){
        get(place)[0][0].toMe(function(e){
          if (replace) {
            e.appendChild(node("div").toMe(function(e){
              e.appendChild(text("error: "));
              e.appendChild(
                node("div", ["style","display: inline-block; vertical-align: top;"]).
                  toMe(function(e){e.innerHTML = body;}));
              e.appendChild(
                node("button",[], 
                  text("Clear")).
                  toMe(function(e){
                    listen(e, "click", function(e){
                      e.target.parentNode.parentNode.removeChild(e.target.parentNode);})}))}))}})});

   // h.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
    h.post(data.forEach(function(v, k){return [k.toString(), encodeURIComponent(v.toString())].join("=")}).join("&"));}

  // init
  listen(window, "load", function init(e){
    action(".resultset", {command: "query"}, true);
    return;})
</script>

    <div class="status">
      <h>Status</h>
      <div class="resultset"></div>
    </div>

<!--#include virtual="ha_BackOffice/ServiceNotifications/bottom.inc"-->
