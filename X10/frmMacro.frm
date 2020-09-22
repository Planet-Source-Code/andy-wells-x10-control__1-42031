VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMacro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Macro"
   ClientHeight    =   3735
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4560
   Icon            =   "frmMacro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   4560
   Begin VB.Timer timCount 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3240
      Top             =   1800
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Options"
      Height          =   1215
      Left            =   1680
      TabIndex        =   10
      Top             =   2400
      Width           =   2775
      Begin VB.HScrollBar scrTime 
         Height          =   255
         LargeChange     =   60
         Left            =   1440
         Max             =   3600
         TabIndex        =   14
         Top             =   720
         Value           =   1
         Width           =   1215
      End
      Begin VB.CheckBox chkCount 
         Caption         =   "&Countdown"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox chkInf 
         Caption         =   "&Infinate"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1440
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "&Run"
      Height          =   375
      Left            =   3240
      TabIndex        =   9
      Top             =   840
      Width           =   1215
   End
   Begin VB.Timer timMacro 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   840
   End
   Begin VB.ListBox lstCommands 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3420
      Left            =   0
      TabIndex        =   8
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add Command"
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   360
      Width           =   1215
   End
   Begin VB.ComboBox cmbCommand 
      Height          =   315
      ItemData        =   "frmMacro.frx":0442
      Left            =   1680
      List            =   "frmMacro.frx":0452
      TabIndex        =   5
      Text            =   "On"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtUnit 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1680
      MaxLength       =   2
      TabIndex        =   2
      Text            =   "1"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtHouse 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1680
      MaxLength       =   1
      TabIndex        =   0
      Text            =   "A"
      Top             =   360
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   0
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblCommand 
      Caption         =   "Command:"
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblUnit 
      Caption         =   "Unit Code:"
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "House Code:"
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
   End
End
Attribute VB_Name = "frmMacro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strHouse As String, strUnit As String, strCommand As String

Private Sub chkCount_Click()
    If blnCount = True Then
        blnCount = False
        scrTime.Enabled = False
    Else
        blnCount = True
        scrTime.Enabled = True
    End If
End Sub

Private Sub chkInf_Click()
    If blnLoop = True Then blnLoop = False Else:  blnLoop = True
End Sub

Private Sub cmdAdd_Click()
    Open strFileName For Append As #1
    Write #1, txtHouse.Text, txtUnit.Text, cmbCommand.Text
    Close #1
    lstCommands.Clear
    Open strFileName For Input As #1
    Do While Not EOF(1)
        Input #1, strHouse, strUnit, strCommand
        lstCommands.AddItem strHouse & ", " & strUnit & ", " & strCommand
    Loop
    Close #1
End Sub

Private Sub cmdClear_Click()
    Dim intMsg As Integer
    intMsg = MsgBox("This will PERMENATELY clear this marco.  Continue?", vbCritical + vbApplicationModal + vbQuestion + vbYesNo, "WARNING!")
    If intMsg = vbNo Then Exit Sub
    Open strFileName For Output As #1
    Close #1
    lstCommands.Clear
End Sub

Public Sub cmdRun_Click()
    If blnCount = True Then
        timCount.Enabled = True
    ElseIf timMacro.Enabled = False Then
        Call Disable
        Open strFileName For Input As #1
        timMacro.Enabled = True
        If blnLoop = True Then cmdRun.Caption = "&Stop"
    Else
        timMacro.Enabled = False
        cmdRun.Caption = "&Run"
        Call Enable
        Close #1
    End If
End Sub

Private Sub Form_Load()
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    Call Disable
    txtHouse.Text = GetSetting(App.Title, "Macro", "HouseCode", "A")
    txtUnit.Text = GetSetting(App.Title, "Macro", "UnitCode", "1")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.Title, "Macro", "HouseCode", txtHouse.Text
    SaveSetting App.Title, "Macro", "UnitCode", txtUnit.Text
    If blnSchedule = False Then frmX10.Show Else: blnSchedule = False
End Sub

Private Sub mnuFileNew_Click()
    On Error GoTo Error
    Dialog.CancelError = True
    Dialog.Flags = cdlOFNFileMustExist
    Dialog.Filter = "X10 Macro Files (*.xmf) | *.xmf"
    Dialog.ShowSave
    strFileName = Dialog.FileName
    Open strFileName For Output As #1
    Close #1
    lstCommands.Clear
    Call Enable
Exit Sub
Error:
End Sub

Private Sub mnuFileOpen_Click()
    On Error GoTo Error
    Dialog.CancelError = True
    Dialog.Flags = cdlOFNFileMustExist
    Dialog.Filter = "X10 Macro Files (*.xmf) | *.xmf"
    Dialog.ShowOpen
    strFileName = Dialog.FileName
    Open strFileName For Input As #1
    lstCommands.Clear
    Do While Not EOF(1)
        Input #1, strHouse, strUnit, strCommand
        lstCommands.AddItem strHouse & ", " & strUnit & ", " & strCommand
    Loop
    Close #1
    Call Enable
Exit Sub
Error:
End Sub

Private Sub scrTime_Change()
    lblTime.Caption = scrTime.Value
    intTime = scrTime.Value
End Sub

Private Sub timCount_Timer()
    If intTime = 0 Then
        timCount.Enabled = False
        chkCount.Value = vbUnchecked
        blnCount = False
        Call cmdRun_Click
        Exit Sub
    End If
    intTime = intTime - 1
    lblTime.Caption = intTime
End Sub

Private Sub timMacro_Timer()
    Static intProgress As Integer
    If EOF(1) Then
        If blnLoop = False Then
            Close #1
            timMacro.Enabled = False
                        Call Enable
            intProgress = 0
            If blnSchedule = False Then
                MsgBox "Macro done!", vbInformation, "Marco"
            Else
                Unload Me
            End If
            Exit Sub
        Else
            Close #1
            intProgress = 0
            Open strFileName For Input As #1
        End If
    End If
    Input #1, strHouse, strUnit, strCommand
    lstCommands.ListIndex = intProgress
    Select Case strCommand
        Case Is = "On"
            frmX10.ControlCm1.Exec strHouse, strUnit, C_ON
        Case Is = "Off"
            frmX10.ControlCm1.Exec strHouse, strUnit, C_OFF
        Case Is = "Bright"
            frmX10.ControlCm1.Exec strHouse, strUnit, C_BRIGHT, 30
        Case Is = "Dim"
            frmX10.ControlCm1.Exec strHouse, strUnit, C_DIM, 30
    End Select
    intProgress = intProgress + 1
End Sub

Private Sub txtHouse_Change()
    txtHouse.Text = UCase(txtHouse.Text)
End Sub

Private Sub txtHouse_GotFocus()
    txtHouse.SelStart = 0
    txtHouse.SelLength = 1
End Sub

Private Sub Enable()
    txtHouse.Enabled = True
    txtUnit.Enabled = True
    cmbCommand.Enabled = True
    cmdAdd.Enabled = True
    cmdClear.Enabled = True
    cmdRun.Enabled = True
    chkInf.Enabled = True
    chkCount.Enabled = True
End Sub

Private Sub Disable()
    txtHouse.Enabled = False
    txtUnit.Enabled = False
    cmbCommand.Enabled = False
    cmdAdd.Enabled = False
    cmdClear.Enabled = False
    If blnLoop = False Then cmdRun.Enabled = False
    chkInf.Enabled = False
    chkCount.Enabled = False
End Sub

Private Sub txtUnit_GotFocus()
    txtUnit.SelStart = 0
    txtUnit.SelLength = Len(txtUnit.Text)
End Sub
