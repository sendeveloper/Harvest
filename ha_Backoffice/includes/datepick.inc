  <script type="text/javascript" src="<%=strPathIncludes%>datepick/jquery-1.9.1.js"></script>
  <script type="text/javascript" src="<%=strPathIncludes%>datepick/jquery-ui.js"></script>
  <script type="text/javascript">
    (function () {
        var elem = document.createElement('input');
        elem.setAttribute('type', 'date');

        if (elem.type === 'text') {
            $('.datepick').datepicker({
                dateFormat: 'yy-mm-dd',
                changeMonth: true,
                changeYear: true,
                showOn: 'button',
                buttonImage: '../images/cal.gif',
                buttonImageOnly: true
            });
            $('.datepick').attr('readonly', 'readonly');
        }
    })(); 
  </script>
