<!--#include virtual="ha_backoffice/includes/Config.asp"-->
<!--#include virtual="includes/connection.asp"-->

<%
  Session("Redirect") = ""
  ColorTab = 3
  PageHeading = "Service Notification Details"
%>

<!--#include virtual="ha_BackOffice/ServiceNotifications/top.inc"-->

<%
    Dim count: count = 0
    Set rs = Sql("select * from casper10.ha_BackOffice.dbo.ha_service_NotificationView order by service, sequence, colid")
%>
<style>
  button.save, button.revert {margin-left: 1em; vertical-align: baseline;}
  h {margin-bottom: 1em; display: inline-block;}
  .modified {border: 1px solid red;}
  .ajax-status {display: inline-block; margin-left: 2em; height: 5em; overflow: scroll; min-width: 45em; max-width: 45em; vertical-align: middle;}

  article {width: 1200px;}
  .details {width: 95%; white-space: normal; margin: 2em; overflow: auto;}
  .details h {font-size: 125%; font-weight: bold;}
  .table h {margin-top: 1em; margin-bottom: 0em; display: block; text-align: center; border-bottom: 1px solid black;}
  div.resultset {margin-top: 1em; margin-bottom: 1em; /*display: inline-block;*/}
  .group div {display: inline-block; margin-left: 1em; margin-right: 1em; vertical-align: top;}
  .group h {text-align: center;}
  table.resultset {margin-left: auto; margin-right: auto; border-spacing: 2em 0em;}
  .resultset tfoot th {text-align: left; padding-top: 1em; font-size: 80%; font-weight: normal;}
  col.Notification-Level td {text-align: right;}
  .resultset {font-size: 80%; white-space: nowrap;}
  .resultset th, .resultset td {font-size: 100%;}
  .resultset thead th {border-bottom: 1px solid black;}

  .console {background-color: black; color: lightgrey; width: 100%; height: auto;}
  .console .table th {color: lightgrey;}


  /* match ha_backoffice style */
    th {text-align: left; border-bottom: 1px solid black; color: black;}
    th.Heading1 {font-size: 14px; text-align: center; border-bottom: 1px solid black; color: black;}

/*
    div.resultset {margin-top: -.5em; border-bottom: 1px solid black; width: 100%;}
*/
    table.resultset {margin-top: 0em; margin-bottom: 1.75em; margin-left: auto; margin-right: auto; width: 100%; border-spacing: .3em;}
    table.resultset > thead > tr:first-child {border-spacing: 0em;}
    table.resultset td, table.resultset th {padding-top: 0.50em; padding-bottom: 0.51em;}
    table.resultset th.service, table.resultset th.notifee {text-align: center; font-size: 90%; padding-top: 0.8em; padding-bottom: .25em; color: black;}

    th {width: 18%; height: auto;}
    th.TableLocation, th.LastRunDuration, th.ActionButton {width: 20%;}
    th.ActionButtons {width: auto;}


</style>
<script type="text/javascript" src="/includes/haFunctions.js"></script>
    <div class="details">
      <div id="ha_service_NotificationsView" class="table">
        <!--h>Services Summary</h-->
        <div id="table"></div>
      </div><!-- ha_service_NotificationsView -->
    </div>

<!--#include virtual="ha_BackOffice/ServiceNotifications/bottom.inc"-->

<script type="text/javascript">
  "use strict";
  node = nodey;

  var query =
    "select [ id]=id, Server='casper10', Operation=Description, [Last Successful]=LastCheckPassed, [Notification Level]=ScheduleStep, [Notification Status] = case when enabled = 0 then 'Disabled' when paused = 1 then 'Paused' else 'Enabled' end " +
    "from casper10.ha_BackOffice.dbo.ha_service_Notifications " +
    "order by server, operation";


  function ajaxError(body) {
    return alert(body);}


  function insertTable(domSelector) {
    return function ajaxCallback(body) {
      return get(domSelector).innerHTML = body;}}

  var ajaxPath = "table-builder.asp";


  Object.prototype.previous = function previous(/* optional count */) {
    var count = ok(arguments[0], 0); //alert([this, count]); count = Math.floor(count);
    for (var node = this; count > 0; --count) {
      node = node.previousSibling;
      //log(textify(node));
      for(; node.nodeType !== 1; node = node.previousSibling) {
        //log("  " + textify(node));
        if (node.previousSibling === null) {return null};}}
    return node;}


  Object.prototype.next = function next(/* optional count */) {
    var count = ok(arguments[0], 0); //alert([this, count]); count = Math.floor(count);
    for (var node = this; count > 0; --count) {
      node = node.nextSibling;
      //log(textify(node));
      for(; node.nodeType !== 1; node = node.nextSibling) {
        //log("  " + textify(node));
        if (node.previousSibling === null) {return null};}}
    return node;}


  Object.prototype.substitute = function substitute(e) {
    this.innerHTML = "";
    this.appendChild(e);
    return this;}


  Object.prototype.remove = function remove() {
    return this.parentNode.removeChild(this);}

  Object.prototype.hide = function hide() {
    return this.style.visibility = "hidden";}

function inspectorGadget(e) {
  e.target.oldstyle = ok(e.target.oldstyle, {});
  e.target.oldstyle.border = e.target.style.border;
  e.target.style.border = "1px solid blue";};

function uninspectorGadget(e) {
  e.target.style = e.target.oldstyle;
  e.target.style.border = e.target.oldstyle.border;
  e.stopPropagation();
  return false;}

// create (nested) nodes -- children are HTML nodes or lists of HTML nodes
function nodey(element, attributes /*children...*/ ) {
  var args = [].slice.call(arguments, 2).reverse();
  var node = document.createElement(element);
  attributes = (!!attributes) ? attributes : {};
  attributes.forEach(function(value, attribute){
    node.setAttribute(attribute, value);})
  while (args.length != 0) {
    var nodelist = args.pop();
    if (nodelist instanceof Array) {
      nodelist.forEach(function(e){node.appendChild(e);})}
    else {
      node.appendChild(nodelist);}}
  nodeEvents.forEach(function(eventHandler){if ((typeof(eventHandler[2] == "undefined") || eventHandler[2](node))) {listen(node, eventHandler[0], eventHandler[1])} return;});
  return node;}


function menu(e) {
  if (e.button !== 2) {return false;}
  var x = xy(e.target)[0];
  var y = xy(e.target)[1];
  nodey("div", {class: "debug-div",
                style: "z-index: 1; border: 1px solid black; width: auto; height: auto; background: white; opacity: 1.00; position: absolute; left: " + x + "px; top: " + y + "px;"},
    nodey("h", {style: "margin-bottom: 0em;"}, text(e.target.nodeName + (e.target.id ? "#" + e.target.id : "") + (e.target.className ? "." + e.target.className : ""))),
    nodey("hr", {style: "margin: 0em;"}),
    nodey("attributes", {}, text("attributes: ")),
    nodey("style", {}, text("style: ")),
    nodey("div", {}, text("parent: " + e.target.parentNode)),
    nodey("div", {}, text("previous: " + e.target.previous())),
    nodey("div", {}, text("next: " + e.target.next())),
    nodey("div", {}, text("children: " + e.target.childNodes))).toMe(function(e){
      get("body")[0][0].appendChild(e);
      listen(e, "mouseout", function(ev){get("body")[0][0].appendChild(nodey("div", {style: "background-color: white; white-space: nowrap"}, text([ev.relatedTarget.innerHTML, ev.target.innerHTML] + " - "), nodey("hr"))); /*e.hide(); */ ev.stopPropagation(); return false}, true);
      return e})};

//insertTable("#table"), ajaxError
  function init(e) {
/*
    get(".table *").forEach(function(e){e.forEach(function(e){
      listen(e, "mouseover", inspectorGadget);
      listen(e, "mousedown", menu);
      listen(e, "mouseout", uninspectorGadget);})});
*/
    ["style#top", "aside"].forEach(function(e){return get(e)[0][0].toMe(function(e){return e.parentNode.removeChild(e)})});
    action("#table", {query: query});
    listen(get("#table thead tr")[0][0].appendChild(node("th", {colspan: "3"}, node("a", {href: "javascript:void(0);"}, text("Add")))),
      "click",
       function(e){
         return openPopUp("popup_sensible.asp?service=")});
    get("#table tbody tr")[0].forEach(function(e){

      listen(
       e.appendChild(node("td", {}, node("a", {href: "javascript:void(0);"}, text("Edit")))),
       "click",
       function(e){
         return openPopUp("popup_sensible.asp?service=" + e.target.parentNode.previous(4).innerHTML);});

      listen(
        e.appendChild(node("td", {}, node("a", {href: "javascript:void(0);"}, text("Pause")))),
        "click",
        function(e){
          e.target.parentNode.previous(2).toMe(function(status){
            status.paused = (e.target.innerHTML == "Pause") ? true : false;
            status.innerHTML = (e.target.parentNode.next(1).childNodes[0].innerHTML == "Enable") ? "Disabled" : status.paused ? "Paused" : "Enabled";
            e.target.innerHTML = e.target.innerHTML.replace(/^Unpause/, "--").replace(/^Pause/, "Unpause").replace(/^--/, "Pause");
            log(textify(e.target.parentNode.parentNode.getAttribute("rel")));
            action(".console", {query: "update sn set Paused = " + (status.paused ? 1 : 0) + "/*select paused*/ from ha_Backoffice.dbo.ha_service_Notifications as sn where convert(nvarchar(max), id) = '" + e.target.parentNode.parentNode.getAttribute("rel") + "'"});});});

      listen(
        e.appendChild(node("td", {}, node("a", {href: "javascript:void(0);"}, text("Disable")))),
        "click",
        function(e){
          e.target.parentNode.previous(3).toMe(function(status){
            status.innerHTML = (e.target.innerHTML == "Disable") ? "Disabled" : status.paused ? "Paused" : "Enabled";
            e.target.innerHTML = e.target.innerHTML.replace(/^Dis/, "--").replace(/^En/, "Dis").replace(/^--/, "En");
            action(".console", {query: "update sn set Enabled = " +  + " select enabled from ha_Backoffice.dbo.ha_service_Notifications where convert(nvarchar(max), id) = '" + e.target.parentNode.parentNode.getAttribute("rel") + "'"});});});

      return;});

    if (get("table.resultset tfoot")[0].length === 0) {
      get("table.resultset")[0][0].appendChild(node("tfoot"));}
    if (get("table.resultset tfoot tr")[0].length === 0) {
      get("table.resultset tfoot")[0][0].appendChild(node("tr"));}
    if (get("table.resultset tfoot tr th")[0].length === 0) {
      get("table.resultset tfoot tr")[0][0].appendChild(node("th"));}
    get("table.resultset tfoot tr th")[0][0].toMe(function(e) {e.setAttribute("colspan", (parseInt(e.getAttribute("colspan")) + 3).toString())});

/*  // this need some handling to prevent mouseout when mousing to children
    get("tfoot th")[0][0].toMe(function(th) {
      var self;
      listen(th, "mouseover", self = function(event) {
        event.stopPropagation();
        event.preventDefault();
        unlisten(th, "mouseover", self);
        var server = th.innerHTML;
        th.substitute(node("select", {}, node("option", {value: server}, text("<cancel>")), ["casper10", "barley1", "barley2", "barley3", "barley4", "philly01", "philly02", "philly03", "philly04", "philly05"].forEach(function(e){
          var selected = (server === e);
          return node("option", server === e, text(e));})));
        listen(th, "mouseout", function(e) {
          e.target.substitute(text(e.target.value));
          listen(th, "mouseover", self);});});});
*/

    return;}

  get("body")[0][0].appendChild(node("div", {class: "console"}));


// this function should be moved to lib.js - <2013-04-04 Thu nathan>
//
// Identify html nodes by type and contents in the output string
function textify(node) {
  return node.toString() + " / " + node.nodeType + " / " + (node.nodeType === 1 ? node.innerHTML : ("\"" + node.nodeValue + "\""));}

/*
//var console = ok(console, ...);
var console = function(e) {alert(e);};
  console.assert = function(condition) {if (!condition) {log("assertion failed: " + condition); throw("assertion failed: " + condition);}},
  console.log = function(e){log("log: " + inspect(e));},
  console.warn = function(e){log("warn: " + inspect(e));},
  console.error = function(e){log("error: " + inspect(e));},
  console.debug = function(e){if (window.debug) alert("debug: " + inspect(e));},
  console.dir = function(e){log("dir: " + inspect(e));},
  console.info = function(e){log("info: " + inspect(e));}

  //console.error(console);
  //console.log("Test: hello.")
*/

var console = get(".console")[0][0];


  listen(window, "load", init);

</script>
