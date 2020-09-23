VERSION 5.00
Begin VB.Form frmLoad 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   255
   LinkTopic       =   "Form1"
   ScaleHeight     =   285
   ScaleWidth      =   255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim z As POINTAPI
If Command$ <> "shortcut" Then Unload Me: Exit Sub
GetCursorPos z
If z.x <> 0 Then Unload Me: Exit Sub
If z.y <> 0 Then Unload Me: Exit Sub
Load frmPW
frmPW.Show
Unload Me
End Sub
