<%
Option Explicit

Const sLocalComputer = "."
Const sWMI_WinSystem = "\root\cimv2"

'Global variables
Dim g_responseText

' *********************************************************
' WMI Helper Functions
'

' Open and return a Connection object with a connection string
Function OpenWMI(strComputer, wmiNamespace)

	Dim objWMIProvider

	On Error Resume Next
    '"winmgmts:{impersonationLevel=impersonate}!\\"
    Set objWMIProvider = GetObject("winmgmts:\\" _
                            & strComputer _
                            & wmiNamespace)

	If Err.Number = 0 Then
		Set OpenWMI = objWMIProvider
	Else
		Set OpenWMI = Nothing
		g_responseText = "WMIError[OpenWMI]: (0x" & Hex(Err.number) & ") " & Err.Description
        Response.AppendToLog "WMIError[OpenWMI]: (0x" & Hex(Err.number) & ") " & Err.Description
	End If

	Set objWMIProvider = Nothing

End Function

Function ExecQuery(wmiObj, strQuery)

    Dim objWMISet

	On Error Resume Next

    Set objWMISet = wmiObj.ExecQuery(strQuery)

	If Err.Number = 0 Then
		Set ExecQuery = objWMISet
	Else
		Set ExecQuery = Nothing
		g_responseText = Err.Description
        Response.AppendToLog "WMIError[ExecQuery]: (0x" & Hex(Err.number) & ") " & Err.Description
	End If

	Set objWMISet = Nothing

End Function

'
' Page Code
'

    Dim strQuery
    Dim objWMI
    Dim objWMICollection
    Dim objEvent
    Dim objWMIDate
    Dim currentDate
    Dim counter
    g_responseText = "OK"

    Do
        'On Error Resume Next

        Set objWMI = OpenWMI(sLocalComputer, sWMI_WinSystem)
        If objWMI Is Nothing Then Exit Do

	Set objWMICollection = objWMI.InstancesOf("Win32_NTLogEventLog")

For Each objEvent In objWMICollection 
     Response.Write( "Filename:           " & objEvent.Log       & vbCrLf & _
		     "Record:             " & objEvent.Record       & vbCrLf)
	Response.Flush
Next


        Set objWMIDate = Server.CreateObject("WbemScripting.SWbemDateTime")
        currentDate = objWMIDate.SetVarDate(DateAdd("h", 1, Now), True)

        'strQuery = "SELECT * FROM Win32_NTLogEvent WHERE Logfile = 'Application' AND SourceName = 'WSH' AND TimeGenerated > '" & objWMIDate.Value & "'"
        strQuery = "SELECT * FROM Win32_NTLogEventLog" 'WHERE Logfile = 'Windows Logs\Application'" 'AND SourceName = '%WSH%'"
        Response.Write(strQuery)
        Set objWMICollection = ExecQuery(objWMI, strQuery)
        If objWMICollection Is Nothing Then Exit Do

        Response.Write("<br/>The query returned (" & objWMICollection.Count & ") records.<br/>")
	'For Each objEvent In objWMICollection
	'    Response.Write("Object: " & objEvent.GetObjectText_) & "<br/>"
	'Next

    Loop While (False)

    Response.Write(g_responseText)

    Set objWMIDate = Nothing
    Set objWMICollection = Nothing
    Set objWMI = Nothing
%>