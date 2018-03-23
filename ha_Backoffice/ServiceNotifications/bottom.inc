  </article>
</div>

  <footer>
<%
  If False Then 'And (Session("loggedin") = True) And (Session("ha_EmpId") = "12") Then
%>
    <div style="background: lightgreen; border: 1px solid black; margin-top: 1.5em;">
<%
    Dim var
    For Each var in Session.Contents
%>
      <label for="session-<%=var%>" style="display: inline-block; text-align: right;"><%=var%>: </label><span id="session-<%=var%>" style="text-align: left; display: inline-block;"><%=Iif((Session(var) = ""), "<em style='opacity: .2'>null</em>", Session(var))%></span><br />
<%
    Next
%>
    </div>
<%
  End If 
%>
  <!--#include virtual="ha_BackOffice/includes/BodyParts/ha_PageTail.inc"-->
  </footer>
</body>
</html>
