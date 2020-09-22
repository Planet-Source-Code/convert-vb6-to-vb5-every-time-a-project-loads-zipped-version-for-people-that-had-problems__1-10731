Attribute VB_Name = "modMain"
Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long
Public Function GetShortPath(strFileName As String) As String
    Dim lngRes As Long, strPath As String
    strPath = String$(165, 0)
    lngRes = GetShortPathName(strFileName, strPath, 164)
    GetShortPath = Left$(strPath, lngRes)
End Function

Public Function GetPathAndFileName(ByVal PathAndFileName, ByRef FileName As String) As String
  Dim lPos As Long
  Dim lLastPos As Long
  
  lPos = InStr(1, PathAndFileName, "\")
  While lPos <> 0
    lLastPos = lPos
    lPos = InStr(lLastPos + 1, PathAndFileName, "\")
  Wend
  
  GetPathAndFileName = Left(PathAndFileName, lLastPos - 1)
  FileName = Mid(PathAndFileName, lLastPos + 1)
  
End Function
Sub Main()
    On Error Resume Next
    Dim property As String
    Dim newfile As String
    Open Command For Input As #1
    Do Until EOF(1)
    Line Input #1, property
    If property = "Retained=0" Then
        Else
        If property = "Retained=1" Then
            Else
            If property = "DebugStartupOption=0" Then
                Else
                If property = "DebugStartupOption=1" Then
                    Else
                 newfile = newfile & property & vbCrLf
             End If
         End If
     End If
    End If
    Loop
    Close #1
    Open Command For Output As #1
    Print #1, newfile
    Close #1
    Dim RetVal
    Dim Path As String
    Dim File As String
    Dim ShortPath
    Dim apppath, cmdline
        If Len(App.Path) <> 2 Then  'if path is not root, add a "\"
        apppath = App.Path & "\"
    Else
        apppath = App.Path
    End If
    Path = GetPathAndFileName(Command, File)
    ShortPath = GetShortPath(Path)
    cmdline = apppath & "Vb5.exe " & ShortPath & "\" & File
    RetVal = Shell(cmdline, vbNormalFocus)
    End
End Sub
