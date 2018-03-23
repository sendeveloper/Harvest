<%
    'Be sure secure path is used
    if Request.ServerVariables("HTTPS") <> "on" then
        response.redirect("https://www.HarvestAmerican.org" & Request.ServerVariables("PATH_INFO"))
    end if
%>