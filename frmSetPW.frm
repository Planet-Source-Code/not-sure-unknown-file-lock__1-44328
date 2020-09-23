VERSION 5.00
Begin VB.Form frmSetPW 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Set Password"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   161
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   2535
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox txtVer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox txtNew 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox txtOld 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Verify Password:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1170
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   " New Password:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Old Password:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1020
   End
End
Attribute VB_Name = "frmSetPW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim code As Long
Dim codeO As Long
Dim codeN As Long
Dim codeV As Long
code = GetSetting(App.Title, "great", "jessus", 152251200)
codeO = Encode(txtOld)
codeN = Encode(txtNew)
codeV = Encode(txtVer)
If codeO <> code Then
    MsgBox "In-Correct old password", vbOKOnly + vbCritical, "Password"
    Exit Sub
End If
If codeN <> codeV Then
    MsgBox "The New and Verify passwords do not match!", vbOKOnly, "Password"
    Exit Sub
End If
If codeO = code Then
    If codeN = codeV Then
        SaveSetting App.Title, "great", "jessus", codeN
        Unload Me
    End If
End If
End Sub

Private Sub Form_Load()
Me.Show vbmoda, frmMain
HideCaret txtOld.hwnd
HideCaret txtNew.hwnd
HideCaret txtVer.hwnd
End Sub

Private Sub txtNew_Change()
HideCaret txtNew.hwnd
End Sub

Private Sub txtNew_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
HideCaret txtNew.hwnd
End Sub

Private Sub txtOld_Change()
HideCaret txtOld.hwnd
End Sub

Private Sub txtOld_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
HideCaret txtOld.hwnd
End Sub

Private Sub txtVer_Change()
HideCaret txtVer.hwnd
End Sub

Private Sub txtVer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
HideCaret txtVer.hwnd
End Sub
