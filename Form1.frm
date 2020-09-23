VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CD-ROM tray example..."
   ClientHeight    =   960
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   960
   ScaleWidth      =   2760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton EjectBtn 
      Caption         =   "&Eject && insert CD-ROM tray..."
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
'   ^ API needed

Private Sub EjectBtn_Click()
    mciSendString "set CDAudio door open", t, 127, 0
    '   ^ open cd rom tray
    mciSendString "set CDAudio door closed", t, 127, 0
    '   ^ close cd rom tray
End Sub

' You can all sorts of fun with this! Just try the FUN.EXE
' for proof!
