VERSION 5.00
Begin VB.Form frmConfirm 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1305
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4590
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
   ScaleHeight     =   1305
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   150
      ScaleHeight     =   945
      ScaleWidth      =   4215
      TabIndex        =   0
      Top             =   150
      Width           =   4275
      Begin VB.CommandButton cmdYes 
         BackColor       =   &H00E0E0E0&
         Cancel          =   -1  'True
         Caption         =   "&Yes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2340
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   540
         Width           =   825
      End
      Begin VB.CommandButton cmdNo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&No"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   540
         Width           =   825
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Remove selected file from Program Launcher?"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   165
         TabIndex        =   1
         Top             =   60
         Width           =   3945
      End
   End
End
Attribute VB_Name = "frmConfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdNo_Click()
    blnRemoveHotKey = False
    Unload Me
    Set frmConfirm = Nothing 'to release the memory
End Sub

Private Sub cmdYes_Click()
    blnRemoveHotKey = True
    Unload Me
    Set frmConfirm = Nothing 'to release the memory
End Sub

Private Sub Form_Load()
    'to make this form always on top
    Call SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    blnRemoveHotKey = False
    prcColorForm Me
End Sub
