VERSION 5.00
Begin VB.Form frmMsg 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1035
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5910
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1035
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   420
      Top             =   330
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   1185
   End
End
Attribute VB_Name = "frmMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    'to make this form always on top
    Call SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Sub

Private Sub Label1_Click()
    Unload Me
End Sub

Private Sub Label1_KeyDown(KeyCode As Integer, Shift As Integer)
    Unload Me
End Sub

Private Sub Timer1_Timer()
    Unload Me
'    Timer1.Enabled = False
End Sub
