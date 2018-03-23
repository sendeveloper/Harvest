<%

    set conn=server.CreateObject("ADODB.Connection")
    conn.Open "driver=SQL Server;server=66.119.55.118,7043;uid=davewj2o;pwd=get2it;database=ha_Backoffice"

    SQL = "SELECT * FROM ha_Documents WHERE [ImageID] = " & Request("ID")
    Set rs = conn.Execute(SQL)

    Response.ContentType = "application/octet-stream"

    ' let the browser know the file name
    Response.AddHeader "Content-Disposition", "attachment;filename=" & Trim(rs("FileName"))

    ' let the browser know the file size
    'Response.AddHeader "Content-Length", rs("FileSize")
    Response.BinaryWrite rs("FileData")
%> 

