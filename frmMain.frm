VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hide [Files - Folders]"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   5280
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   369
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   352
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timRefresh 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   1320
      Top             =   2400
   End
   Begin VB.Frame Frame1 
      Caption         =   "Rename:"
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   4080
      Width           =   5055
      Begin VB.Frame Frame3 
         Height          =   855
         Left            =   1320
         TabIndex        =   7
         Top             =   240
         Width           =   3495
         Begin VB.CommandButton cmdShow 
            Caption         =   "Un-Lock"
            Height          =   495
            Left            =   1800
            TabIndex        =   9
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton cmdHide 
            Caption         =   "Lock"
            Height          =   495
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame2 
         Height          =   855
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   975
         Begin VB.OptionButton optFolder 
            Caption         =   "Folder"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton optFile 
            Caption         =   "File"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   480
            Width           =   735
         End
      End
   End
   Begin VB.FileListBox File1 
      Height          =   3990
      Hidden          =   -1  'True
      Left            =   2520
      TabIndex        =   2
      Top             =   0
      Width           =   2775
   End
   Begin VB.DirListBox Dir1 
      Height          =   3690
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2415
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileSet 
         Caption         =   "&Set Password"
         Shortcut        =   ^S
      End
      Begin VB.Menu d1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub HideFolder()
Dim msg As String
msg = Dir1.List(Dir1.ListIndex)
msg = "ren " & msg & " " & RemovePath(msg) & Chr$(254)
Open App.path & "\temp.bat" For Output As #1
    Print #1, msg
Close #1
ShellExecute GetDesktopWindow(), "open", App.path & "\temp.bat", vbNullString, vbNullString, 0
timRefresh = True
End Sub

Public Sub ShowFolder()
Dim msg As String
msg = Dir1.List(Dir1.ListIndex)
If Right$(msg, 1) <> "_" Then Exit Sub
msg = Left$(msg, Len(msg) - 1)
msg = "ren " & msg & Chr$(254) & " " & RemovePath(msg)
Open App.path & "\temp.bat" For Output As #1
    Print #1, msg
Close #1
ShellExecute GetDesktopWindow(), "open", App.path & "\temp.bat", vbNullString, vbNullString, 0
timRefresh = True
End Sub

Public Sub HideFile()
Dim msg As String
Dim path As String
Dim file As String
Dim ext As String
path = GetPath(Dir1.path)
file = GetFile(File1.FileName)
ext = GetExt(File1.FileName)
msg = "ren " & path & file & ext & " " & file & Chr$(254) & ext
Open App.path & "\temp.bat" For Output As #1
    Print #1, msg
Close #1
ShellExecute GetDesktopWindow(), "open", App.path & "\temp.bat", vbNullString, vbNullString, 0
timRefresh = True
End Sub

Public Sub ShowFile()
Dim msg As String
Dim path As String
Dim file As String
Dim ext As String
path = GetPath(Dir1.path)
file = GetFile(File1.FileName)
file = Left$(file, Len(file) - 1)
ext = GetExt(File1.FileName)
msg = "ren " & path & file & Chr$(254) & ext & " " & file & ext
Open App.path & "\temp.bat" For Output As #1
    Print #1, msg
Close #1
ShellExecute GetDesktopWindow(), "open", App.path & "\temp.bat", vbNullString, vbNullString, 0
timRefresh = True
End Sub

Private Sub cmdHide_Click()
If optFolder = True Then
    HideFolder
End If
If optFile = True Then
    HideFile
End If
End Sub

Private Sub cmdShow_Click()
If optFolder = True Then
    ShowFolder
End If
If optFile = True Then
    ShowFile
End If
End Sub

Private Sub Dir1_Change()
On Error GoTo skip
File1.path = Dir1.path
Exit Sub
skip:
File1.path = Drive1.Drive
End Sub

Private Sub Dir1_GotFocus()
If optFolder.Value = False Then optFolder.Value = True
End Sub

Private Sub Drive1_Change()
Dir1.path = Drive1.Drive
End Sub

Private Sub File1_GotFocus()
If optFile.Value = False Then optFile.Value = True
End Sub

Private Sub mnuFileExit_Click()
Unload Me
End Sub

Private Sub mnuFileSet_Click()
Load frmSetPW
End Sub

Private Sub timRefresh_Timer()
Dir1.Refresh
File1.Refresh
Kill App.path & "\temp.bat"
timRefresh = False
End Sub
