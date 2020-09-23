VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Program Launcher"
   ClientHeight    =   3945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6390
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   6390
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2640
      Top             =   1770
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Height          =   3705
      Left            =   150
      ScaleHeight     =   3645
      ScaleWidth      =   6105
      TabIndex        =   15
      Top             =   150
      Width           =   6165
      Begin VB.OptionButton optPredefined 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Predefined"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4890
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   2370
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.OptionButton optUser 
         Caption         =   "User Defined"
         Height          =   285
         Left            =   3750
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   2370
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox chkWin 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Win"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1140
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2760
         Width           =   675
      End
      Begin VB.CommandButton cmdFolder 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   345
         Left            =   540
         MaskColor       =   &H80000005&
         Picture         =   "frmMain.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   " Select target folder... "
         Top             =   2730
         Width           =   345
      End
      Begin VB.TextBox txtFileName 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   210
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   2400
         Width           =   5685
      End
      Begin VB.CommandButton cmdFile 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   345
         Left            =   150
         MaskColor       =   &H80000005&
         Picture         =   "frmMain.frx":7EEC
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   " Select target file... "
         Top             =   2730
         Width           =   345
      End
      Begin VB.CommandButton cmdHelp 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Help"
         Height          =   315
         Left            =   4080
         MaskColor       =   &H80000005&
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   " About Program Launcher, Usage and Error Lookup "
         Top             =   3240
         Width           =   945
      End
      Begin VB.CommandButton cmdLaunch 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Launch"
         Height          =   315
         Left            =   2100
         MaskColor       =   &H80000005&
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   " Launch selected file "
         Top             =   3240
         Width           =   945
      End
      Begin MSComctlLib.ListView lstvwPrograms 
         Height          =   2085
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   3678
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   16777215
         BackColor       =   16711680
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Exit"
         Height          =   315
         Left            =   5070
         MaskColor       =   &H80000005&
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   " Stop using Program Launcher "
         Top             =   3240
         Width           =   945
      End
      Begin VB.PictureBox picTray 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   5430
         Picture         =   "frmMain.frx":9BE6
         ScaleHeight     =   480
         ScaleWidth      =   450
         TabIndex        =   19
         Top             =   2730
         Width           =   450
      End
      Begin VB.CommandButton cmdHide 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Cancel          =   -1  'True
         Caption         =   "Hi&de"
         Height          =   315
         Left            =   3090
         MaskColor       =   &H80000005&
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   " Hide Program Launcher (Press Win+A to make visible) "
         Top             =   3240
         Width           =   945
      End
      Begin VB.CommandButton cmdDelete 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Remove"
         Height          =   315
         Left            =   1110
         MaskColor       =   &H80000005&
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   " Delete selected shortcut "
         Top             =   3240
         Width           =   945
      End
      Begin VB.CheckBox chkCtrl 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ctrl"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2010
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2760
         Width           =   675
      End
      Begin VB.CheckBox chkShift 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Shift"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2760
         Width           =   675
      End
      Begin VB.CheckBox chkAlt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Alt"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3750
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2760
         Width           =   675
      End
      Begin VB.ComboBox cmbKey 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4620
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2760
         Width           =   735
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   5010
         Top             =   300
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdSave 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Save"
         Height          =   315
         Left            =   120
         MaskColor       =   &H80000005&
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   " Save shortcut "
         Top             =   3240
         Width           =   945
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1860
         TabIndex        =   24
         Top             =   2820
         Width           =   120
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2730
         TabIndex        =   18
         Top             =   2820
         Width           =   120
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   4470
         TabIndex        =   17
         Top             =   2820
         Width           =   120
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3600
         TabIndex        =   16
         Top             =   2820
         Width           =   120
      End
   End
   Begin VB.Label Label9 
      BackColor       =   &H00404040&
      Height          =   15
      Left            =   0
      TabIndex        =   23
      Top             =   4005
      Width           =   6825
   End
   Begin VB.Label Label8 
      BackColor       =   &H00E0E0E0&
      Height          =   15
      Left            =   -60
      TabIndex        =   22
      Top             =   0
      Width           =   6735
   End
   Begin VB.Label Label7 
      BackColor       =   &H00404040&
      Height          =   4215
      Left            =   6465
      TabIndex        =   21
      Top             =   -30
      Width           =   15
   End
   Begin VB.Label Label6 
      BackColor       =   &H00E0E0E0&
      Height          =   4005
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   15
   End
   Begin VB.Menu mnuMenus 
      Caption         =   "Main Menus"
      Visible         =   0   'False
      Begin VB.Menu mnuProgList 
         Caption         =   "Show Program List"
      End
      Begin VB.Menu mnuActivate 
         Caption         =   "Activate when Windows starts"
      End
      Begin VB.Menu d1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim blnDrag As Boolean
Dim lTop As Long, lLeft As Long
Dim intIndex As Long
Dim strF As String

Private Sub chkCtrl_Click()
    chkCtrl.FontBold = chkCtrl.Value
End Sub

Private Sub chkShift_Click()
    chkShift.FontBold = chkShift.Value
End Sub

Private Sub chkAlt_Click()
    chkAlt.FontBold = chkAlt.Value
End Sub

Private Sub chkWin_Click()
    chkWin.FontBold = chkWin.Value
End Sub

Private Sub cmdFile_Click()
On Error GoTo Err1
    CommonDialog1.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
    CommonDialog1.CancelError = True
    CommonDialog1.Filter = "All files (*.*)|*.*" 'displays all files
    CommonDialog1.DialogTitle = "Select Target file"
    CommonDialog1.ShowOpen
    txtFileName.Text = CommonDialog1.FileName
    strF = "I"
    Exit Sub
Err1:
End Sub

Public Sub cmdClose_Click()
    prcCloseProgram
End Sub

Private Sub cmdDelete_Click()
Dim intRecNo As Integer
Dim lRet As Long
    
    Picture1.SetFocus
    frmConfirm.Show vbModal
    If blnRemoveHotKey = False Then Exit Sub
    intRecNo = 1
    
    lRet = UnregisterHotKey(frmIcon.hwnd, lstvwPrograms.SelectedItem.Index)
    Kill strDataFile 'delete data file
    lstvwPrograms.ListItems.Remove lstvwPrograms.SelectedItem.Index 'remove selected item
    prcWrite 're-write the data file
    prcUnregisterAll
    prcRegisterAll
End Sub

Private Sub cmdHelp_Click()
    Picture1.SetFocus
    frmHelp.Show vbModal
End Sub

Private Sub cmdHide_Click()
    Picture1.SetFocus
    Unload Me
    Set frmMain = Nothing
End Sub

Private Sub cmdLaunch_Click()
    Picture1.SetFocus
    If Not lstvwPrograms.SelectedItem Is Nothing Then
        prcLaunch lstvwPrograms.SelectedItem.ListSubItems(1).Text & "\" & lstvwPrograms.SelectedItem.Text
    End If
End Sub

Private Sub cmdSave_Click()
Dim lRet As Long
Dim vModifierKey
Dim lstItem As ListItem
Dim bSave As Boolean
Dim sKeys As String
    
    Picture1.SetFocus
    If txtFileName.Text = "" Then
        prcShowMsg "Select Target File name"
        cmdFile_Click
        Exit Sub
    ElseIf Dir(txtFileName.Text, vbDirectory + vbNormal) = "" Then
        prcShowMsg "Target file/folder not found" '
        Exit Sub
    ElseIf chkWin.Value = 0 And chkAlt.Value = 0 And chkShift.Value = 0 And chkCtrl.Value = 0 Then
        prcShowMsg "Select a system key"
        Exit Sub
    ElseIf cmbKey.Text = "" Then
        prcShowMsg "Select a key"
        cmbKey.SetFocus
        Exit Sub
    End If
    
    If chkWin.Value = 1 Then sKeys = "Win"
    If chkCtrl.Value = 1 Then sKeys = IIf(sKeys = "", "Ctrl", sKeys & "+Ctrl")
    If chkShift.Value = 1 Then sKeys = IIf(sKeys = "", "Shift", sKeys & "+Shift")
    If chkAlt.Value = 1 Then sKeys = IIf(sKeys = "", "Alt", sKeys & "+Alt")
    sKeys = sKeys & "+" & cmbKey.Text

    With lstvwPrograms.ListItems
        For lRet = 1 To .Count
            vModifierKey = UCase(.Item(lRet).ListSubItems(2).Text)
            If UCase(sKeys) = vModifierKey Then
                prcShowMsg "Selected combination of keys is already assigned to " & vbCrLf & "'" & .Item(lRet).ListSubItems(1).Text & "\" & .Item(lRet).Text & "'"
                Exit Sub
            End If
        Next
    End With
    
    vModifierKey = ""
    If chkWin.Value = 1 Then vModifierKey = MOD_WIN
    If chkCtrl.Value = 1 Then vModifierKey = vModifierKey + MOD_CONTROL
    If chkShift.Value = 1 Then vModifierKey = vModifierKey + MOD_SHIFT
    If chkAlt.Value = 1 Then vModifierKey = vModifierKey + MOD_ALT
    
    With objPLKeys
        .pKeyWin = IIf(chkWin.Value, MOD_WIN, 0)
        .pKeyAlt = IIf(chkAlt.Value, MOD_ALT, 0)
        .pKeyCtrl = IIf(chkCtrl.Value, MOD_CONTROL, 0)
        .pKeyShift = IIf(chkShift.Value, MOD_SHIFT, 0)
        .pKey = Asc(cmbKey.Text)
        .pTargetFile = txtFileName.Text
        
        lRet = RegisterHotKey(frmIcon.hwnd, intHotKeyCount + 1, vModifierKey, Asc(cmbKey.Text))
    End With
    vModifierKey = ""
    If lRet = 1 Then
        With objPLKeys
            Set lstItem = lstvwPrograms.ListItems.Add(Text:=CommonDialog1.FileTitle) ', Index:=.pHotKeyID
            lstItem.ListSubItems.Add Text:=fnFilefolder(txtFileName.Text, 2)
            lstItem.ListSubItems.Add Text:=sKeys
        End With
        prcWrite
    Else
        frmMsg.Label1.BackColor = vbRed
        prcShowMsg "Unable to save hotkey. GetLastError returned: " & GetLastError
    End If
    chkWin.Value = 0
    chkAlt.Value = 0
    chkCtrl.Value = 0
    chkShift.Value = 0
    txtFileName.Text = ""
End Sub

Private Sub cmdFolder_Click()
Dim myShell As New Shell
Dim myFolder As Folder
Set myFolder = myShell.BrowseForFolder(Me.hwnd, "Select folder", 16)
If Not myFolder Is Nothing Then
    txtFileName.Text = myFolder.Items.Item.Path
End If
CommonDialog1.FileName = ""
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    blnDrag = True
    lTop = Y
    lLeft = X
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If blnDrag Then
        Me.Top = Me.Top + (Y - lTop)
        Me.Left = Me.Left + (X - lLeft)
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    blnDrag = False
End Sub

Private Sub Form_Load()
Dim intOldMode As Integer

    'to make this form always on top
    Call SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    
    '**** TO HIDE THE TITLE BAR ******
    SetWindowLong frmMain.hwnd, GWL_STYLE, GetWindowLong(frmMain.hwnd, GWL_STYLE) And Not WS_CAPTION
    intOldMode = frmMain.ScaleMode
    frmMain.ScaleMode = 1
    frmMain.Height = frmMain.Height - 300
    frmMain.ScaleMode = intOldMode
    '*********************************
    lstvwPrograms.ColumnHeaders.Add Text:="File/Folder", Width:=1340
    lstvwPrograms.ColumnHeaders.Add Text:="Path", Width:=3000
    lstvwPrograms.ColumnHeaders.Add Text:="Shortcut key", Width:=1530
    
    With cmbKey
        .AddItem ""
        For intIndex = 65 To 90
            .AddItem Chr(intIndex)
        Next
        For intIndex = 48 To 57
            .AddItem Chr(intIndex)
        Next
    End With
    intIndex = 0
    prcLoadData
    prcColorForm Me
    
    Me.Left = Screen.Width - Me.Width
    Me.Top = Screen.Height - Me.Height
    Me.Refresh
    Me.Visible = True
End Sub

Private Sub prcLoadData()
'this procedure is to load the hotkey details from data.dat into listview
Dim lstItem As ListItem
Dim strText As String

    lstvwPrograms.ListItems.Clear
    intIndex = 1
    Open strDataFile For Random Access Read As #1 'binary file access
    With objPLKeys
    Do
        Get #1, intIndex, objPLKeys  'read from file. read one record at a time
        If .pKey = 0 Then Exit Do 'there are no more records in file
        
        strText = fnFilefolder(.pTargetFile, 1)
        Set lstItem = lstvwPrograms.ListItems.Add(Text:=strText) ', Index:=.pHotKeyID
        
        strText = fnFilefolder(.pTargetFile, 2)
        lstItem.ListSubItems.Add Text:=strText
        strText = ""
        If .pKeyWin = MOD_WIN Then
            strText = "Win"
        End If
        If .pKeyCtrl = MOD_CONTROL Then
            strText = strText & IIf(strText = "", "Ctrl", "+Ctrl")
        End If
        If .pKeyShift = MOD_SHIFT Then
            strText = strText & IIf(strText = "", "Shift", "+Shift")
        End If
        If .pKeyAlt = MOD_ALT Then
            strText = strText & IIf(strText = "", "Alt", "+Alt")
        End If
        strText = strText & "+" & Chr(.pKey)
        lstItem.ListSubItems.Add Text:=strText
        intIndex = intIndex + 1
    Loop While True
    End With
    Close #1
End Sub

Private Sub prcWrite()
Dim intRecNo As Integer
Dim strKeys As String
    intRecNo = 1
    'to read record by record from binary file, use only structures.
    'use an instance of a structure to write to file.
    Open strDataFile For Random As #1 'binary file access
    With lstvwPrograms.ListItems
    intHotKeyCount = 0
    For intIndex = 1 To .Count
        strKeys = UCase(.Item(intIndex).ListSubItems(2).Text) 'hotkeys available in listview
        
        With objPLKeys
'            .pHotKeyID = intIndex
            If InStr(1, strKeys, "WIN", vbTextCompare) <> 0 Then .pKeyWin = MOD_WIN Else .pKeyWin = 0
            If InStr(1, strKeys, "CTRL", vbTextCompare) <> 0 Then .pKeyCtrl = MOD_CONTROL Else .pKeyCtrl = 0
            If InStr(1, strKeys, "SHIFT", vbTextCompare) <> 0 Then .pKeyShift = MOD_SHIFT Else .pKeyShift = 0
            If InStr(1, strKeys, "ALT", vbTextCompare) <> 0 Then .pKeyAlt = MOD_ALT Else .pKeyAlt = 0
            .pKey = Asc(Right(strKeys, 1))
            With lstvwPrograms.ListItems.Item(intIndex)
                objPLKeys.pTargetFile = .ListSubItems(1).Text & IIf(.Text = "", "", "\" & .Text)
            End With
        End With
        
        Put #1, intIndex, objPLKeys 'write struct to file
        intHotKeyCount = intHotKeyCount + 1
    Next
    End With
    Close #1
End Sub

Private Sub optUser_Click()
    Picture1.SetFocus
End Sub

Private Sub optPredefined_Click()
    Picture1.SetFocus
End Sub

'to retrieve file name or folder from complete path and filename
Private Function fnFilefolder(sz As String, i As Integer) As String
Dim strText As String
Dim intSlashPos As Long
    strText = StrReverse(Trim(sz))
    If InStr(1, strText, ".", vbTextCompare) = 0 Then
        'string does not have '.', its folder name
        If i = 1 Then
            strText = ""
        Else
            strText = sz
        End If
    Else
        'string does have '.', its file name
        'folder name also have have '.', in that case,
        'it wud be difficult to seperate path and file name
        'so, that cannot be handled
        
        intSlashPos = InStr(1, strText, "\", vbTextCompare)
        If i = 1 Then
            'return path
            strText = StrReverse(Left(strText, intSlashPos - 1))
        ElseIf i = 2 Then
            'return file name
            strText = StrReverse(Mid(strText, intSlashPos + 1))
        End If
    End If
    fnFilefolder = strText
End Function
