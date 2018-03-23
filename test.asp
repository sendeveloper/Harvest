<%

	'dim server
	aserver = "db.Zip2Tax.com"
    dbUsername = "z2t_link"
    dbPassword = "H^2p6~r"
    dbName = "zip2tax"
    connString = "driver=SQL Server;Server=" & aserver & "; Database=" & dbName & "; User Id=" & dbUsername & "; password=" & dbPassword & ";"
		
	response.write connString

        'Using cn As New SqlConnection(connString)
		
		dim connString
		Set connString = Server.CreateObject("ADODB.Connection")
		'connString.Open
		
  Dim connCasper10: Set connCasper10 = Server.CreateObject("ADODB.Connection")
  connCasper10.Open "driver=SQL Server;server=db.Zip2Tax.com;uid=z2t_link;pwd=H^2p6~r"		
%>