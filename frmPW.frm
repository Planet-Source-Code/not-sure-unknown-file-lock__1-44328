VERSION 5.00
Begin VB.Form frmPW 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Password"
   ClientHeight    =   510
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   34
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   161
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPW 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E24E82&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmPW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Show
HideCaret txtPW.hwnd
End Sub

Private Sub txtPW_Change()
HideCaret txtPW.hwnd
End Sub

Private Sub txtPW_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    ProcessCode
End If
End Sub

Public Sub ProcessCode()
Dim code As Long, moveon As Boolean
moveon = False
code = GetSetting(App.Title, "great", "jessus", -1)
If code = -1 Then
    If Encode(txtPW) = 152251200 Then
        moveon = True
    Else
        moveon = False
    End If
Else
    If Encode(txtPW) = code Then
        moveon = True
    Else
        moveon = False
    End If
End If
If moveon = False Then Unload Me: Exit Sub
If moveon = True Then
    Load frmMain
    frmMain.Show
    Unload Me
End If
End Sub
