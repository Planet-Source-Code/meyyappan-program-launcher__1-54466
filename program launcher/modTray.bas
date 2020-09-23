Attribute VB_Name = "modTray"
Option Explicit

'to store left and top positoin of windows
Private Type POINTAPI
    X As Long
    Y As Long
End Type

'windows message structure
Public Type Msg
    hwnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

'trayicon structure
Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

'to store window position
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'hotkey constants
Public Const MOD_ALT = &H1
Public Const MOD_CONTROL = &H2
Public Const MOD_SHIFT = &H4
Public Const MOD_WIN = &H8
Public Const PM_REMOVE = &H1
Public Const WM_HOTKEY = &H312

'trayicon constants
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

'trayicon callback constants
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205

'shellexecute error constants
Public Const ERROR_FILE_NOT_FOUND = 2&
Public Const ERROR_PATH_NOT_FOUND = 3&
Public Const ERROR_BAD_FORMAT = 11&
Public Const SE_ERR_ACCESSDENIED = 5            '  access denied
Public Const SE_ERR_ASSOCINCOMPLETE = 27
Public Const SE_ERR_DDEBUSY = 30
Public Const SE_ERR_DDEFAIL = 29
Public Const SE_ERR_DDETIMEOUT = 28
Public Const SE_ERR_DLLNOTFOUND = 32
Public Const SE_ERR_FNF = 2                     '  file not found
Public Const SE_ERR_NOASSOC = 31
Public Const SE_ERR_PNF = 3                     '  path not found
Public Const SE_ERR_OOM = 8                     '  out of memory
Public Const SE_ERR_SHARE = 26
Public Const GWL_STYLE = (-16)
Public Const WS_CAPTION = &HC00000
'Public Const sBarWidth = 262

'setwindowpos constants
Public Const HWND_TOPMOST& = -1
Public Const SWP_NOMOVE& = &H2
Public Const SWP_NOSIZE& = &H1

'for making desktop icons transparent
Const LVM_FIRST = &H1000
Const LVM_SETTEXTBKCOLOR = (LVM_FIRST + 38)
Const CLR_NONE = &HFFFFFFFF
Const SWP_HIDEWINDOW = &H80
Const SWP_SHOWWINDOW = &H40

'API declarations
Public Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Declare Function Shell_NotifyIcon Lib "Shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetWindowPos& Lib "user32" (ByVal hwnd&, ByVal hWndInsertAfter&, ByVal X&, ByVal Y&, ByVal cx&, ByVal cy&, ByVal wFlags&)
Public Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Public Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long) As Long
Public Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As Msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Public Declare Function WaitMessage Lib "user32" () As Long
Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'user defined structure; used to store keys
Public Type PLFileStruct
    pKeyWin As Long
    pKeyShift As Long
    pKeyCtrl As Long
    pKeyAlt As Long
    pKey As Long
    pTargetFile As String
    'pHotKeyID As Long
End Type
Public strDataFile As String
Public objPLKeys As PLFileStruct
Public intHotKeyCount As Integer
Public blnRemoveHotKey As Boolean
'Public blnDontProcessMessage As Boolean

Dim TrayI As NOTIFYICONDATA
Dim taskBarPos As RECT
Dim hTaskBar As Long
    
Public Sub prcPutMeOnTray()
'this procedure is to add an icon to the system tray
    TrayI.cbSize = Len(TrayI)
    
    'Link the trayicon to this picturebox
    'Right click or Double click on the system tray icon will
    'trigger the picTray's events
    TrayI.hwnd = frmIcon.picTray.hwnd
    
    TrayI.uId = 1&
    TrayI.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    
    'raise picTray's event when right click or double click on system tray icon
    TrayI.ucallbackMessage = WM_RBUTTONDOWN
    
    TrayI.hIcon = frmIcon.picTray.Picture
    TrayI.szTip = "Program Launcher" & Chr$(0)
    Shell_NotifyIcon NIM_ADD, TrayI
End Sub

'to remove the icon from system tray
Public Sub prcRemoveMeFromTray()
    Shell_NotifyIcon NIM_DELETE, TrayI
End Sub

Public Function fnGetFormPosition() As RECT
    hTaskBar = FindWindow("Shell_TrayWnd", vbNullString)
    GetWindowRect hTaskBar, taskBarPos
End Function

'to display any message
Public Sub prcShowMsg(strMsg As String)
On Error Resume Next
    With frmMsg
        Load frmMsg
        .Label1.Caption = strMsg
        .Label1.Height = 100 + .TextHeight(strMsg)
        .Label1.Width = 240 + (.TextWidth(strMsg))
        .Height = 300 + .Label1.Height
        .Width = 330 + .Label1.Width
        .Timer1.Interval = Len(.Label1.Caption) * 50
        .Timer1.Enabled = True
        prcColorForm frmMsg
        .Show vbModal
        .Timer1.Enabled = False
    End With
End Sub

Public Sub Main()
Dim lng1 As Long


    strDataFile = App.Path & "\data.dat"
    
    'In Win98 all appliations and processes appear in Task Manager list.
    'So, "App.TaskVisible = False" does not have an effect in Win98.
    'In Win2000 (or higher), Applications and Processes are listed seperately.
    'This statement will hide the application from Task Manager list.
    'However it appears under "Processes".
    App.TaskVisible = False
    'Load frmMsg
    
    'you can use Spy++ from Microsoft Visual Studio Tool to find the class name of an exe
    lng1 = FindWindow("ThunderRT6FormDC", "Program Launcher")
    If lng1 <> 0 Then
        frmMsg.Label1.BackColor = vbRed
        prcShowMsg "Program Launcher is already active"
        SetFocus lng1
        End
    End If
    Load frmIcon
End Sub

'to launch the file/folder
Public Sub prcLaunch(FileName As String)
Dim lng1 As Long
    
    lng1 = ShellExecute(frmMsg.hwnd, "open", FileName, "", 0, 10)
    frmMsg.Label1.BackColor = vbRed
    Select Case lng1
    Case ERROR_BAD_FORMAT:
        prcShowMsg "'" & FileName & "'" & vbCrLf & "Bad file format"
    Case ERROR_FILE_NOT_FOUND:
        prcShowMsg "'" & FileName & "'" & vbCrLf & "File not found"
    Case ERROR_PATH_NOT_FOUND:
        prcShowMsg "'" & FileName & "'" & vbCrLf & "Path not found"
    Case SE_ERR_ACCESSDENIED:
        prcShowMsg "'" & FileName & "'" & vbCrLf & "File access denied"
    Case SE_ERR_ASSOCINCOMPLETE:
        prcShowMsg "'" & FileName & "'" & vbCrLf & "File association information not complete"
    Case SE_ERR_DDEBUSY:
        prcShowMsg "'" & FileName & "'" & vbCrLf & "DDE operation is busy"
    Case SE_ERR_DDEFAIL:
        prcShowMsg "'" & FileName & "'" & vbCrLf & "DDE operation failed"
    Case SE_ERR_DDETIMEOUT:
        prcShowMsg "'" & FileName & "'" & vbCrLf & "DDE operation timed out"
    Case SE_ERR_DLLNOTFOUND:
        prcShowMsg "'" & FileName & "'" & vbCrLf & "Dynamic-link library not found"
    Case SE_ERR_FNF:
        prcShowMsg "'" & FileName & "'" & vbCrLf & "File not found"
    Case SE_ERR_PNF:
        prcShowMsg "'" & FileName & "'" & vbCrLf & "Path not found"
    Case SE_ERR_NOASSOC:
        prcShowMsg "'" & FileName & "'" & vbCrLf & "No program is associated with this file"
    Case SE_ERR_OOM:
        prcShowMsg "'" & FileName & "'" & vbCrLf & "Out of memory"
    Case SE_ERR_SHARE:
        prcShowMsg "'" & FileName & "'" & vbCrLf & "File is already open. Cannot share"
    End Select
    frmMsg.Label1.BackColor = vbBlue
End Sub

'to draw colored lines in the form
Public Sub prcColorForm(myForm As Form)
Dim intIndex As Integer
    With myForm
        .DrawStyle = vbInsideSolid
        .DrawMode = vbCopyPen
        .ScaleMode = vbTwips
        .DrawWidth = 2
        .ScaleHeight = 200
        .AutoRedraw = True
    End With
    For intIndex = 1 To 100
        myForm.Line (0, intIndex)-(Screen.Width, intIndex - 1), RGB(155 + intIndex, 155 + intIndex, 155 + intIndex), B
    Next intIndex
    For intIndex = 101 To 200
        myForm.Line (0, intIndex)-(Screen.Width, intIndex - 1), RGB(255 + (101 - intIndex), 255 + (101 - intIndex), 255 + (101 - intIndex)), B
    Next intIndex
End Sub

Public Sub prcCloseProgram()
    prcUnregisterAll
    prcRemoveMeFromTray
    Unload frmMsg
    Unload frmMain
    Unload frmHelp
    Unload frmConfirm
    Unload frmIcon
    
    Set frmMsg = Nothing
    Set frmMain = Nothing
    Set frmHelp = Nothing
    Set frmConfirm = Nothing
    Set frmIcon = Nothing
    prcShowMsg "Program Launcher closed"
    End
End Sub

'to process all the messages
Public Sub prcProcessMessage()
    Dim Message As Msg
    Do
        WaitMessage 'waits for a message
        
        'check whether the message corresponds to the registered hotkey
        If PeekMessage(Message, frmIcon.hwnd, WM_HOTKEY, WM_HOTKEY, PM_REMOVE) Then
            If Message.wParam = 101 Then
                frmMain.Show
            ElseIf Message.wParam = 100 Then
                'to lock the pc running on Win 2000
                Shell "rundll32 user32.dll,LockWorkStation", vbNormalFocus
            Else
                Open strDataFile For Random Access Read As #1 'binary file access
                Get #1, Message.wParam, objPLKeys 'read from file. read one record at a time
                prcLaunch objPLKeys.pTargetFile
                Close #1
            End If
        End If
        DoEvents
    Loop While True
End Sub

'to open text file resister hot keys for programs that were already assigned hotkeys
Public Sub prcRegisterAll()
Dim lRet As Long
Dim intIndex As Integer
    Open strDataFile For Random Access Read As #1 'binary file access
    Do
        Get #1, intIndex + 1, objPLKeys 'read from file. read one record at a time
        If objPLKeys.pKey = 0 Then
            Close #1
            If intIndex = 0 Then frmMain.Show
            prcProcessMessage
            Exit Do
        Else
            With objPLKeys
            UnregisterHotKey frmIcon.hwnd, intIndex + 1 '.pHotKeyID
            lRet = RegisterHotKey(frmIcon.hwnd, intIndex + 1, .pKeyWin + .pKeyCtrl + .pKeyShift + .pKeyAlt, .pKey)
            End With
            If lRet = 0 Then
                frmMsg.Label1.BackColor = vbRed
                prcShowMsg "Unable to save hotkey. GetLastError returned: " & GetLastError
                Close #1
                prcProcessMessage
                Exit Sub
            End If
        End If
        intIndex = intIndex + 1
    Loop While True
End Sub

'to unregister all hotkeys
Public Sub prcUnregisterAll()
Dim i As Integer
    For i = 1 To intHotKeyCount
        UnregisterHotKey frmIcon.hwnd, i
    Next
    UnregisterHotKey frmIcon.hwnd, 100
    UnregisterHotKey frmIcon.hwnd, 101
End Sub
