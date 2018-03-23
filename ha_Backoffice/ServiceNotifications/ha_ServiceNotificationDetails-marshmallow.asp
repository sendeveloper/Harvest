<%
  Session("Redirect") = ""
  ColorTab = 3
  PageHeading = "Service Notification Details"
%>

<!--#include virtual="ha_BackOffice/ServiceNotifications/top.inc"-->

<% 
    Dim count
    count = 0
    Set rs = sql("select * from casper10.ha_BackOffice.dbo.ha_service_NotificationView order by service, sequence, colid")
%>
<style>
  button.save, button.revert {margin-left: 1em; vertical-align: baseline;}
  h {margin-bottom: 1em; display: inline-block;}
  .modified {border: 1px solid red;}
  .ajax-status {display: inline-block; margin-left: 2em; height: 5em; overflow: scroll; min-width: 45em; max-width: 45em; vertical-align: middle;}

  .details {width: 1200px; white-space: normal;}
  .table h {margin-top: 1em; margin-bottom: 0em; font-size: 80%; display: block; text-align: left;}
  div.resultset {margin-top: 1em; margin-bottom: 1em; /*display: inline-block;*/}
  .group div {display: inline-block; margin-left: 1em; margin-right: 1em; vertical-align: top;}
  .group h {text-align: center;}
</style>

    <div class="details">
      <h>Details</h>
      <div id="ha_service_NotificationsView" class="table">
        <h>ha_service_NotificationsView</h>
        <%=sqlTable("select * from casper10.ha_BackOffice.dbo.ha_service_NotificationView order by service, sequence, colid")%>
      </div><!-- ha_service_NotificationsView -->
      <div id="ha_service_NotificationSchedule" class="table">
        <h>ha_service_NotificationSchedule</h>
        <%=sqlTable("select * from casper10.ha_BackOffice.dbo.ha_service_NotificationSchedule order by serviceid, sequence")%>
      </div><!-- ha_service_NotificationSchedule -->
      <div id="ha_service_NotificationLevels" class="table">
        <h>ha_service_NotificationLevels</h>
        <%=sqlTable("select nl.id, ns.Sequence, nl.colid, nl.Description, nl.Enabled from casper10.ha_BackOffice.dbo.ha_service_NotificationLevels as nl left join casper10.ha_BackOffice.dbo.ha_service_NotificationSchedule as ns on ns.id = nl.ScheduleId order by colid, ns.Sequence ")%>
      </div><!-- ha_service_NotificationLevels -->
    </div>

<!--#include virtual="ha_BackOffice/ServiceNotifications/bottom.inc"-->
