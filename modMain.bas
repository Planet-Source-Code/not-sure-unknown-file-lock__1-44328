Attribute VB_Name = "modMain"
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function HideCaret Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long

Public Type POINTAPI
        x As Long
        y As Long
End Type

Public Function Encode(code As String) As Long
Dim i As Integer
Dim cout As Long
cout = 1
For i = 1 To Len(code)
    cout = cout * Asc(Mid$(code, i, 1))
Next i
Encode = cout
End Function

Public Function RemovePath(msg As String) As String
Dim i As Integer
i = InStr(1, msg, "\")
While i <> 0
    msg = Right$(msg, Len(msg) - i)
    i = InStr(1, msg, "\")
Wend
RemovePath = msg
End Function

Public Function GetPath(path As String) As String
If Right$(path, 1) <> "\" Then path = path & "\"
GetPath = path
End Function

Public Function GetFile(file As String) As String
Dim i As Integer
i = InStr(1, file, ".")
If i <> 0 Then
    file = Left$(file, i - 1)
End If
GetFile = file
End Function

Public Function GetExt(ext As String) As String
Dim i As Integer
i = InStr(1, ext, ".")
If i <> 0 Then
    ext = Right$(ext, Len(ext) - i + 1)
End If
If i = 0 Then ext = ""
GetExt = ext
End Function
