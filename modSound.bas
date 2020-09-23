Attribute VB_Name = "modSound"
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Dim tmp As String * 255
Dim ShortPath As Long
Dim ShortPathAndFie As String

Public Sub PlayMP3(Filename As String)
    ShortPath = GetShortPathName(Filename, tmp, 255)
    ShortPathAndFie = Left$(tmp, ShortPath)
    mciSendString "Open " & ShortPathAndFie & " Alias MM", 0, 0, 0
    mciSendString "Play MM", 0, 0, 0
End Sub


Public Sub PauseMP3()
    mciSendString "Stop MM", 0, 0, 0
End Sub


Public Sub StopMP3()
    mciSendString "Stop MM", 0, 0, 0
    mciSendString "Close MM", 0, 0, 0
End Sub
