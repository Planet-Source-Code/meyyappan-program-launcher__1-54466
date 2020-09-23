VERSION 5.00
Begin VB.Form frmHelp 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4005
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6480
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   6480
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
      Height          =   3705
      Left            =   150
      ScaleHeight     =   3645
      ScaleWidth      =   6105
      TabIndex        =   0
      Top             =   150
      Width           =   6165
      Begin VB.OptionButton optError 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Errors"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1980
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3240
         Width           =   825
      End
      Begin VB.CommandButton cdmClose 
         BackColor       =   &H00E0E0E0&
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   285
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3240
         Width           =   825
      End
      Begin VB.OptionButton optUsage 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Usage"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1050
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3240
         Width           =   825
      End
      Begin VB.OptionButton optAbout 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&About"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   3240
         Width           =   825
      End
      Begin VB.TextBox txtMessage 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   3015
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   120
         Width           =   5865
      End
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cdmClose_Click()
    Unload Me
    Set frmHelp = Nothing
End Sub

Private Sub Form_Activate()
Picture1.SetFocus
End Sub

Private Sub Form_Load()
    'to make this form always on top
    Call SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    optAbout.Value = True
    prcColorForm Me
End Sub

Private Sub optAbout_Click()
    If Me.Visible Then Picture1.SetFocus
    optAbout.FontBold = True
    optUsage.FontBold = False
    optError.FontBold = False
    txtMessage.Text = vbTab & "Program Launcher can be used to launch files/programs quickly by pressing a combination of keys or on click of a button." & vbCrLf & vbCrLf & vbTab & "Just select the file/program you want to launch and assign a shortcut-key. Pressing this keys will launch the file/program. This is similar to using shortcut-key for a shortcut file (link file). The difference is, only one instance of the file/program will be launched with using the shortcut keys of a shortcut file. You may open as many instances as you want with Program Launcher." & vbCrLf & vbCrLf & vbTab & " Program Launcher should be running in hidden mode to launch the file/program on pressing the shortcut-keys. When Program Launcher is running, a keyboard icon is present near the system clock." & vbCrLf & vbCrLf & "NOTE: Shortcut keys does not work when Program Launcher is visible."
End Sub

Private Sub optAbout_GotFocus()
    Picture1.SetFocus
End Sub

Private Sub optUsage_GotFocus()
    Picture1.SetFocus
End Sub

Private Sub optError_GotFocus()
    Picture1.SetFocus
End Sub

Private Sub optUsage_Click()
    Picture1.SetFocus
    optAbout.FontBold = False
    optUsage.FontBold = True
    optError.FontBold = False
    With txtMessage
    .Text = "To add file/program to the list:" & vbCrLf & "1. Select the file/program." & vbCrLf & "2. Assign atleast one system key (Ctrl/Alt/Shift) and a character." & vbCrLf & "3. Click on Save."
    .Text = .Text & vbCrLf & vbCrLf & "To remove file/program from the list:" & vbCrLf & "1. Select the file/program." & vbCrLf & "2. Click in Remove."
    .Text = .Text & vbCrLf & vbCrLf & "To launch file/program:" & vbCrLf & "Press the assigned shortcut keys." & vbCrLf & "or" & vbCrLf & "1. Select the file/program in the list." & vbCrLf & "2. Click on Launch."
    .Text = .Text & vbCrLf & vbCrLf & "To hide Program Launcher:" & vbCrLf & "Click on Hide button."
    .Text = .Text & vbCrLf & vbCrLf & "To show Program Launcher:" & vbCrLf & "Double-click on the keyboard icon near the system tray." & vbCrLf & "or" & vbCrLf & "Right click on the keyboard icon and select 'Show Program List'." & vbCrLf & "or" & vbCrLf & "Press Ctrl+F12."
    .Text = .Text & vbCrLf & vbCrLf & "To add Program Launcher to Windows startup:" & vbCrLf & "Right click on the keyboard icon and ensure that 'Activate when Windows starts' is selected."
    .Text = .Text & vbCrLf & vbCrLf & "To remove Program Launcher from Windows startup:" & vbCrLf & "Right click on the keyboard icon and ensure that 'Activate when Windows starts' is unselected."
    End With
End Sub

Private Sub optError_Click()
    Picture1.SetFocus
    optAbout.FontBold = False
    optUsage.FontBold = False
    optError.FontBold = True
    With txtMessage
    .Text = vbTab & "When launching any file/program there are possibility of errors like 'file is already open','access denied','out of memory',etc. Given below is the list of errors that may occur in Program Launcher (Error messages are displayed in red color background)."
    .Text = .Text & vbCrLf & "1. Bad file format"
    .Text = .Text & vbCrLf & "2. File not found"
    .Text = .Text & vbCrLf & "3. Path not found"
    .Text = .Text & vbCrLf & "4. File access denied"
    .Text = .Text & vbCrLf & "5. Out of memory"
    .Text = .Text & vbCrLf & "6. No program is associated with this file"
    .Text = .Text & vbCrLf & "7. File association information not complete"
    .Text = .Text & vbCrLf & "8. DDE operation is busy"
    .Text = .Text & vbCrLf & "9. DDE operation failed"
    .Text = .Text & vbCrLf & "10.DDE operation timed out"
    .Text = .Text & vbCrLf & "11.Dynamic-link library not found"
    End With
End Sub
