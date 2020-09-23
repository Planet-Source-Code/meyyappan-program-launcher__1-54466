VERSION 5.00
Begin VB.Form frmIcon 
   Caption         =   "Program Launcher"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1860
      Top             =   1230
   End
   Begin VB.PictureBox picTray 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   180
      Picture         =   "frmIcon.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   180
      Width           =   480
   End
   Begin VB.Menu mnuMenus 
      Caption         =   "Main Menus"
      Begin VB.Menu mnuProgList 
         Caption         =   "Show &Program List"
      End
      Begin VB.Menu mnuActivate 
         Caption         =   "Activate when Windows starts"
      End
      Begin VB.Menu d1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intIndex As Long

'this form will not be displayed. this is just to create the system tray icon and
'display a menu on right click
Private Sub Form_Load()

    'saves in System Registry
    'HKEY_CURRENT_USER\Software\VB and VBA Program Settings\ProgramLauncher\Settings
    If GetSetting(App.EXEName, "Settings", "Startup", "1") = "1" Then
        mnuActivate.Checked = True
    Else
        mnuActivate.Checked = False
    End If
    mnuProgList.Enabled = False
    prcPutMeOnTray
    prcActivate mnuActivate.Checked
    
    'to listen to keystores globally
    prcShowMsg "Program Launcher is activated"
    RegisterHotKey frmIcon.hwnd, 100, MOD_WIN, vbKeySpace
    RegisterHotKey frmIcon.hwnd, 101, MOD_WIN, Asc("A")
    mnuProgList.Enabled = True
    prcRegisterAll  'to register hotkeys that were already saved
End Sub

Private Sub mnuActivate_Click()
    If mnuActivate.Checked Then
        mnuActivate.Checked = False
        prcActivate False
    Else
        mnuActivate.Checked = True
        prcActivate True
    End If
End Sub

Private Sub mnuExit_Click()
    prcCloseProgram
End Sub

Private Sub mnuProgList_Click()
    If frmMsg.Visible = False Then
        frmMain.Show
    End If
End Sub

Private Sub picTray_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim Msg As Long
    Msg = X / Screen.TwipsPerPixelX
    If Msg = WM_LBUTTONDBLCLK Then
        'on system tray icon double click, show the main screen
        Load frmMain
        BringWindowToTop frmMain.hwnd
        frmMain.SetFocus
    ElseIf Msg = WM_RBUTTONUP Then
        'on system tray icon right click, show menu
        Me.PopupMenu mnuMenus
    End If
End Sub

Private Sub prcActivate(AutoActivate As Boolean)
'this procedure is to do the setting of Start when Windows start
Dim WshShell As Object, oShellLink As Object

    Set WshShell = CreateObject("Wscript.Shell")
    If AutoActivate Then
        
        'saves in System Registry
        'HKEY_CURRENT_USER\Software\VB and VBA Program Settings\ProgramLauncher\Settings
        SaveSetting App.EXEName, "Settings", "Startup", "1"
        
        If Dir(WshShell.specialfolders("startup") & "\Program Launcher.lnk") = "" Then
        'to create a shortcut in Starup folder
            Set oShellLink = WshShell.CreateShortcut(WshShell.specialfolders("startup") & "\Program Launcher.lnk")
            oShellLink.TargetPath = App.Path & "\Program Launcher.exe"
            oShellLink.Save
        End If
    Else
        
        'saves in System Registry
        'HKEY_CURRENT_USER\Software\VB and VBA Program Settings\ProgramLauncher\Settings
        SaveSetting App.EXEName, "Settings", "Startup", "0"
        
        If Dir(WshShell.specialfolders("startup") & "\Program Launcher.lnk") <> "" Then
        'to delete the shortcut file in Startup folder
            DeleteFile WshShell.specialfolders("startup") & "\Program Launcher.lnk"
        End If
    End If
End Sub
