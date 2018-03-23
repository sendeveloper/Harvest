<%
'item: a string -- a collection key
'collection: 
'
' Returns: True if `item' is in collection; False otherwise.

Function member(item, cCollection)
        member = False
        
        On Error Resume Next
        If cCollection(item) Is Not Nothing Then
                member = True
        End If
End Function

Function list(aarray)
        Set list = Server.CreateObject("Scripting.Dictionary")
        For Each i in aarray
                list.Add i,i
        Next        
End Function

Function format(obj)
  Dim str 
  Dim i
  Dim space

  space = ""
  'str = TypeName(obj) & ": "
  str = ""

  Select Case TypeName(obj)
    Case "Recordset"
      str = str & "<"
      For Each i in obj.fields
        str = str & space & obj(i)
        space = " "
      Next
      str = str & ">"
    Case "Field"
     str = str & "'" & obj.Name & "'"
    Case "Integer", "Long"
      str = str & cstr(obj)
    Case "String"
      str = chr(34) & obj & chr(34)
    Case "Variant()"
      str = str & "["
      For Each i In obj
        str = str + space + format(i)
        space = " "
      Next 
      str = str & "]"
    Case "Dictionary"
      str = str & "("
      For Each i In obj
        'str = str & space & "(" & format(i) & " . " & format(obj(i)) & ")"
        str = str & space & format(obj(i))
        space = " "
      Next
      str = str & ")"
    Case Else
      str = str & "&lt;? " & TypeName(obj)& " ?&gt;"
  End Select
  format = str
End Function

Select Case Request("persist")
  Case "1"
    Session("DebuggingPersists") = 1
  Case "0"
    Session("DebuggingPersists") = 0
   Case Else
    Session("DebuggingPersists") = Session("DebuggingPersists")
End Select

If Session("DebuggingPersists") = 1 Then
   Session("Debugging") = Session("Debugging")
ElseIf (Session("z2t_loggedin") > "") And (Session("z2t_UserName") > "") Then
    Response.Write format((list(Array("nathan", "testy"))))
    If member(Session("z2t_UserName"), list(Array("nathan", "testy"))) Then
      If (Request("debug") = "1") Then
        Session("Debugging") = "otia"
      ElseIf (Request("debug") = "0") Then
        Session("Debugging") = ""
      Else
        If (Request ("debug") > "") Then
          Session("Debugging") = Request("Debug")
        Else
          Session("Debugging") = " "
        End If
        If InStr(Request("Debug"), "u") Then
          Session("NewUser") = "True"
        End If
      End If
    Else
      'Other users don't need to debug
      Session("Debugging") = ""
    End If
ElseIf Request("debug") = ".mkujoim.m" Then
  Session("Debugging") = " "
Else
        'Anonymous users don't need to debug
        Session("Debugging") = ""
End If


%>
