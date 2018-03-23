<%
  Session("Redirect") = ""
  ColorTab = 3
  PageHeading = "Service Notification Edit"
%>

<!--#include virtual="ha_BackOffice/ServiceNotifications/top.inc"-->
<script>
  var changes = [];

  function debug(/* args */) {
    var args = [].slice.call(arguments);
    return alert("debug: {" + args.forEach(function(e){return ok(e, "<undefined>").toString()}).join(", ") + "}");}
  

  ajaxPath = "/ha_BackOffice/ServiceNotifications/ha_ServiceNotificationAjax.asp";

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
          if (replace | !replace) {
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


  listen(window, "load", function init(e){
    get(".resultset td").forEach(function(e){e.forEach(function(e){
      e.undo = e.innerHTML;
      listen(e, "dblclick", function(e){
        e.target.edititng = true;
        e.target.toMe(function(e){
          var content = e.innerHTML;
          e.childNodes.forEach(function(e){
            if (typeof(e.parentNode) === "undefined") {return;} // webkit bug
            e.parentNode.removeChild(e)});
          e.appendChild(node("input", ["type","text", "value",content, "size",content.length, "style","width: auto;"]).toMe(function(e){
                          listen(e, "change", function(e){
                            debug("--" + e);
                            debug(e, e.parentNode);
                            e.target.parentNode.edited = true;
                            classify(e.target.parentNode, "modified");
                            if (!e.target.parentNode.changeNode) {
                              e.target.parentNode.changeNode = {node: e.target.parentNode,
                                                                undo: []};
                              changes.push(e.target.parentNode.changeNode);}
                            e.target.parentNode.changeNode.toMe(function(changeNode){
                              changeNode.undo.push(e.target.parentNode.undo);});
                            e.target.parentNode.innerHTML = e.target.parentNode.undo = e.target.value;});
                          listen(e, "blur", function(e){
                              if (!e.target.parentNode.edited) {
                                e.target.parentNode.innerHTML = e.target.value;}});}));
          e.edited = false;
          e.childNodes[0].select();})});})});

    listen(get("button.save")[0][0], "click", function(e){
      var modifiedAnything = false;
      get(".modified").forEach(function(e){e.forEach(function(e){
        modifiedAnything = true;
        message("acting: " + e);
        declassify(e, "modified");
        action(".ajax-status", {command: "save", 
                                field: e.className,
                                type: e.getAttribute("rel"),
                                value: e.innerHTML,
                                table: e.parentNode.parentNode.parentNode.parentNode.parentNode.id.split(".").last(), 
                                where: "id=" + e.parentNode.querySelector("td").innerHTML});
        e.undo = e.innerHTML;
        e.changeNode = null;});
      message(modifiedAnything ? "Changes saved." : "No changes to save.");})
      changes = [];});
    listen(get("button.revert")[0][0], "click", function(e){
      var modifiedAnything = false;
      changes.forEach(function(e){
        modifiedAnything = true;
        e.node.undo = e.undo[0];
        e.node.innerHTML = e.undo[0];
        e.node.changeNode = null;
        declassify(e.node, "modified");});
      message(modifiedAnything ? "Discarded changes." : "No changes to discard.");
      changes = [];});});

  function message(messageText) {
    get(".ajax-status")[0][0].toMe(function(e){
      e.appendChild(node("div", [], text(messageText)));
      e.scrollTop = e.scrollHeight;})
    return;}

  function last() {
    return this[this.length - 1];}
  Object.prototype.last = last;
</script>
<style>
  button.save, button.revert {margin-left: 1em; vertical-align: baseline;}
  h {margin-bottom: 1em; display: inline-block;}
  .modified {border: 1px solid red;}
  .ajax-status {display: inline-block; margin-left: 2em; height: 5em; overflow: scroll; min-width: 45em; max-width: 45em; vertical-align: middle;}

  .edit {width: 1200px; white-space: normal;}
  .table h {margin-top: 1em; margin-bottom: 0em; font-size: 80%; display: block; text-align: left;}
  div.resultset {margin-top: 1em; margin-bottom: 1em; /*display: inline-block;*/}
  .group div {display: inline-block; margin-left: 1em; margin-right: 1em; vertical-align: top;}
  .group h {text-align: center;}
</style>
<%
  sqlFieldTypes = True
  Set rs = sql("select * from dallas01.ha_BackOffice.dbo.ha_service_NotificationView order by service, sequence, colid")
%>
    <div class="edit">
      <h>Edit</h><button class="save">save</button><button class="revert">revert</button><div class="ajax-status"></div><hr>
<%
  Dim tables: tables = "ha_BackOffice.dbo.ha_service_NotificationView/" &_
                       "ha_BackOffice.dbo.ha_service_NotificationServicesView/" &_
                       "ha_BackOffice.dbo.ha_service_NotificationSchedule;serviceid,sequence/" &_
                       "ha_BackOffice.dbo.ha_service_NotificationLevels;scheduleid,colid/" &_
                       "ha_BackOffice.dbo.ha_service_notifees;Enabled desc/" &_
                       "ha_BackOffice.dbo.ha_service_Notifications;Priority"
  Call ShowTables(tables)
%>

        <div class="group">
<%

  tables = "dallas01.ha_BackOffice.dbo.ha_service_JobTypes;id/" &_
           "dallas01.ha_BackOffice.dbo.ha_service_NotificationMethods;id"
  Call ShowTables(tables)
%>
          <div id="ha_empAccounts" class="table">
            <h>ha_empAccouns</h>
<%
  Call SqlTableInsert("select id, EmpId, Firstname, LastName from dallas01.ha_BackOffice.dbo.ha_EmpAccounts where ActiveStatus = 1 order by FirstName, LastName")
%>
          </div>
        </div>
    </div>

<!--#include virtual="ha_BackOffice/ServiceNotifications/bottom.inc"-->

<%
' tables is a slash-separated string of semicolon-separated pairs table;orderby
Function ShowTables(tables)
  For each TableString in Split(tables, "/")
     Dim TablePair: TablePair = Split(TableString, ";")
     Dim table: table = TablePair(0)
     Dim orderby: 
     If Ubound(TablePair) < 1 Then
       orderby = ""
     Else
       orderby = " order by " & TablePair(1)
     End If
      
     Response.Write("<div id=""" & table & """ class=""table"">")
     Response.Write("<h>" & LastDot(table) & "</h>")
     Call sqlTableInsert("select * from " & table & orderby)
     Response.Write("</div>")
  Next
End Function

Function LastDot(text)
  Dim parts: parts = Split(text, ".")
  LastDot = parts(Ubound(parts))
End Function


Function present(stream, e) 
   stream.addChild()
End Function   
%>
