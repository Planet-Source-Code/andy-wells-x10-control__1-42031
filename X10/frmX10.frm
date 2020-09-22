VERSION 5.00
Object = "{7D7F2F33-1D39-11D3-AA96-006097C0E8C9}#5.1#0"; "cm17a.ocx"
Begin VB.Form frmX10 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "X10 Control"
   ClientHeight    =   6585
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   3345
   Icon            =   "frmX10.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   3345
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraControls 
      Caption         =   "Remote Interface"
      Height          =   6015
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   3135
      Begin VB.CommandButton cmdOff 
         Caption         =   "Off"
         Height          =   255
         Index           =   1
         Left            =   1005
         TabIndex        =   67
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdOn 
         Caption         =   "On"
         Height          =   255
         Index           =   1
         Left            =   405
         TabIndex        =   66
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdOff 
         Caption         =   "Off"
         Height          =   255
         Index           =   2
         Left            =   1005
         TabIndex        =   65
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton cmdOn 
         Caption         =   "On"
         Height          =   255
         Index           =   2
         Left            =   405
         TabIndex        =   64
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton cmdOff 
         Caption         =   "Off"
         Height          =   255
         Index           =   3
         Left            =   1005
         TabIndex        =   63
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton cmdOn 
         Caption         =   "On"
         Height          =   255
         Index           =   3
         Left            =   405
         TabIndex        =   62
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton cmdOff 
         Caption         =   "Off"
         Height          =   255
         Index           =   4
         Left            =   1005
         TabIndex        =   61
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton cmdOn 
         Caption         =   "On"
         Height          =   255
         Index           =   4
         Left            =   405
         TabIndex        =   60
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton cmdOff 
         Caption         =   "Off"
         Height          =   255
         Index           =   5
         Left            =   1005
         TabIndex        =   59
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton cmdOn 
         Caption         =   "On"
         Height          =   255
         Index           =   5
         Left            =   405
         TabIndex        =   58
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton cmdOff 
         Caption         =   "Off"
         Height          =   255
         Index           =   6
         Left            =   1005
         TabIndex        =   57
         Top             =   2040
         Width           =   615
      End
      Begin VB.CommandButton cmdOn 
         Caption         =   "On"
         Height          =   255
         Index           =   6
         Left            =   405
         TabIndex        =   56
         Top             =   2040
         Width           =   615
      End
      Begin VB.CommandButton cmdOff 
         Caption         =   "Off"
         Height          =   255
         Index           =   7
         Left            =   1005
         TabIndex        =   55
         Top             =   2400
         Width           =   615
      End
      Begin VB.CommandButton cmdOn 
         Caption         =   "On"
         Height          =   255
         Index           =   7
         Left            =   405
         TabIndex        =   54
         Top             =   2400
         Width           =   615
      End
      Begin VB.CommandButton cmdOff 
         Caption         =   "Off"
         Height          =   255
         Index           =   8
         Left            =   1005
         TabIndex        =   53
         Top             =   2760
         Width           =   615
      End
      Begin VB.CommandButton cmdOn 
         Caption         =   "On"
         Height          =   255
         Index           =   8
         Left            =   405
         TabIndex        =   52
         Top             =   2760
         Width           =   615
      End
      Begin VB.CommandButton cmdOff 
         Caption         =   "Off"
         Height          =   255
         Index           =   9
         Left            =   1005
         TabIndex        =   51
         Top             =   3120
         Width           =   615
      End
      Begin VB.CommandButton cmdOn 
         Caption         =   "On"
         Height          =   255
         Index           =   9
         Left            =   405
         TabIndex        =   50
         Top             =   3120
         Width           =   615
      End
      Begin VB.CommandButton cmdOff 
         Caption         =   "Off"
         Height          =   255
         Index           =   10
         Left            =   1005
         TabIndex        =   49
         Top             =   3480
         Width           =   615
      End
      Begin VB.CommandButton cmdOn 
         Caption         =   "On"
         Height          =   255
         Index           =   10
         Left            =   405
         TabIndex        =   48
         Top             =   3480
         Width           =   615
      End
      Begin VB.CommandButton cmdOff 
         Caption         =   "Off"
         Height          =   255
         Index           =   11
         Left            =   1005
         TabIndex        =   47
         Top             =   3840
         Width           =   615
      End
      Begin VB.CommandButton cmdOn 
         Caption         =   "On"
         Height          =   255
         Index           =   11
         Left            =   405
         TabIndex        =   46
         Top             =   3840
         Width           =   615
      End
      Begin VB.CommandButton cmdOff 
         Caption         =   "Off"
         Height          =   255
         Index           =   12
         Left            =   1005
         TabIndex        =   45
         Top             =   4200
         Width           =   615
      End
      Begin VB.CommandButton cmdOn 
         Caption         =   "On"
         Height          =   255
         Index           =   12
         Left            =   405
         TabIndex        =   44
         Top             =   4200
         Width           =   615
      End
      Begin VB.CommandButton cmdOff 
         Caption         =   "Off"
         Height          =   255
         Index           =   13
         Left            =   1005
         TabIndex        =   43
         Top             =   4560
         Width           =   615
      End
      Begin VB.CommandButton cmdOn 
         Caption         =   "On"
         Height          =   255
         Index           =   13
         Left            =   405
         TabIndex        =   42
         Top             =   4560
         Width           =   615
      End
      Begin VB.CommandButton cmdOff 
         Caption         =   "Off"
         Height          =   255
         Index           =   14
         Left            =   1005
         TabIndex        =   41
         Top             =   4920
         Width           =   615
      End
      Begin VB.CommandButton cmdOn 
         Caption         =   "On"
         Height          =   255
         Index           =   14
         Left            =   405
         TabIndex        =   40
         Top             =   4920
         Width           =   615
      End
      Begin VB.CommandButton cmdOff 
         Caption         =   "Off"
         Height          =   255
         Index           =   15
         Left            =   1005
         TabIndex        =   39
         Top             =   5280
         Width           =   615
      End
      Begin VB.CommandButton cmdOn 
         Caption         =   "On"
         Height          =   255
         Index           =   15
         Left            =   405
         TabIndex        =   38
         Top             =   5280
         Width           =   615
      End
      Begin VB.CommandButton cmdOff 
         Caption         =   "Off"
         Height          =   255
         Index           =   16
         Left            =   1005
         TabIndex        =   37
         Top             =   5640
         Width           =   615
      End
      Begin VB.CommandButton cmdOn 
         Caption         =   "On"
         Height          =   255
         Index           =   16
         Left            =   405
         TabIndex        =   36
         Top             =   5640
         Width           =   615
      End
      Begin VB.CommandButton cmdDim 
         Caption         =   "Dim"
         Height          =   255
         Index           =   1
         Left            =   2325
         TabIndex        =   35
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdBright 
         Caption         =   "Bright"
         Height          =   255
         Index           =   1
         Left            =   1725
         TabIndex        =   34
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdDim 
         Caption         =   "Dim"
         Height          =   255
         Index           =   2
         Left            =   2325
         TabIndex        =   33
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton cmdBright 
         Caption         =   "Bright"
         Height          =   255
         Index           =   2
         Left            =   1725
         TabIndex        =   32
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton cmdDim 
         Caption         =   "Dim"
         Height          =   255
         Index           =   3
         Left            =   2325
         TabIndex        =   31
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton cmdBright 
         Caption         =   "Bright"
         Height          =   255
         Index           =   3
         Left            =   1725
         TabIndex        =   30
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton cmdDim 
         Caption         =   "Dim"
         Height          =   255
         Index           =   4
         Left            =   2325
         TabIndex        =   29
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton cmdBright 
         Caption         =   "Bright"
         Height          =   255
         Index           =   4
         Left            =   1725
         TabIndex        =   28
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton cmdDim 
         Caption         =   "Dim"
         Height          =   255
         Index           =   5
         Left            =   2325
         TabIndex        =   27
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton cmdBright 
         Caption         =   "Bright"
         Height          =   255
         Index           =   5
         Left            =   1725
         TabIndex        =   26
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton cmdDim 
         Caption         =   "Dim"
         Height          =   255
         Index           =   6
         Left            =   2325
         TabIndex        =   25
         Top             =   2040
         Width           =   615
      End
      Begin VB.CommandButton cmdBright 
         Caption         =   "Bright"
         Height          =   255
         Index           =   6
         Left            =   1725
         TabIndex        =   24
         Top             =   2040
         Width           =   615
      End
      Begin VB.CommandButton cmdDim 
         Caption         =   "Dim"
         Height          =   255
         Index           =   7
         Left            =   2325
         TabIndex        =   23
         Top             =   2400
         Width           =   615
      End
      Begin VB.CommandButton cmdBright 
         Caption         =   "Bright"
         Height          =   255
         Index           =   7
         Left            =   1725
         TabIndex        =   22
         Top             =   2400
         Width           =   615
      End
      Begin VB.CommandButton cmdDim 
         Caption         =   "Dim"
         Height          =   255
         Index           =   8
         Left            =   2325
         TabIndex        =   21
         Top             =   2760
         Width           =   615
      End
      Begin VB.CommandButton cmdBright 
         Caption         =   "Bright"
         Height          =   255
         Index           =   8
         Left            =   1725
         TabIndex        =   20
         Top             =   2760
         Width           =   615
      End
      Begin VB.CommandButton cmdDim 
         Caption         =   "Dim"
         Height          =   255
         Index           =   9
         Left            =   2325
         TabIndex        =   19
         Top             =   3120
         Width           =   615
      End
      Begin VB.CommandButton cmdBright 
         Caption         =   "Bright"
         Height          =   255
         Index           =   9
         Left            =   1725
         TabIndex        =   18
         Top             =   3120
         Width           =   615
      End
      Begin VB.CommandButton cmdDim 
         Caption         =   "Dim"
         Height          =   255
         Index           =   10
         Left            =   2325
         TabIndex        =   17
         Top             =   3480
         Width           =   615
      End
      Begin VB.CommandButton cmdBright 
         Caption         =   "Bright"
         Height          =   255
         Index           =   10
         Left            =   1725
         TabIndex        =   16
         Top             =   3480
         Width           =   615
      End
      Begin VB.CommandButton cmdDim 
         Caption         =   "Dim"
         Height          =   255
         Index           =   11
         Left            =   2325
         TabIndex        =   15
         Top             =   3840
         Width           =   615
      End
      Begin VB.CommandButton cmdBright 
         Caption         =   "Bright"
         Height          =   255
         Index           =   11
         Left            =   1725
         TabIndex        =   14
         Top             =   3840
         Width           =   615
      End
      Begin VB.CommandButton cmdDim 
         Caption         =   "Dim"
         Height          =   255
         Index           =   12
         Left            =   2325
         TabIndex        =   13
         Top             =   4200
         Width           =   615
      End
      Begin VB.CommandButton cmdBright 
         Caption         =   "Bright"
         Height          =   255
         Index           =   12
         Left            =   1725
         TabIndex        =   12
         Top             =   4200
         Width           =   615
      End
      Begin VB.CommandButton cmdDim 
         Caption         =   "Dim"
         Height          =   255
         Index           =   13
         Left            =   2325
         TabIndex        =   11
         Top             =   4560
         Width           =   615
      End
      Begin VB.CommandButton cmdBright 
         Caption         =   "Bright"
         Height          =   255
         Index           =   13
         Left            =   1725
         TabIndex        =   10
         Top             =   4560
         Width           =   615
      End
      Begin VB.CommandButton cmdDim 
         Caption         =   "Dim"
         Height          =   255
         Index           =   14
         Left            =   2325
         TabIndex        =   9
         Top             =   4920
         Width           =   615
      End
      Begin VB.CommandButton cmdBright 
         Caption         =   "Bright"
         Height          =   255
         Index           =   14
         Left            =   1725
         TabIndex        =   8
         Top             =   4920
         Width           =   615
      End
      Begin VB.CommandButton cmdDim 
         Caption         =   "Dim"
         Height          =   255
         Index           =   15
         Left            =   2325
         TabIndex        =   7
         Top             =   5280
         Width           =   615
      End
      Begin VB.CommandButton cmdBright 
         Caption         =   "Bright"
         Height          =   255
         Index           =   15
         Left            =   1725
         TabIndex        =   6
         Top             =   5280
         Width           =   615
      End
      Begin VB.CommandButton cmdDim 
         Caption         =   "Dim"
         Height          =   255
         Index           =   16
         Left            =   2325
         TabIndex        =   5
         Top             =   5640
         Width           =   615
      End
      Begin VB.CommandButton cmdBright 
         Caption         =   "Bright"
         Height          =   255
         Index           =   16
         Left            =   1725
         TabIndex        =   4
         Top             =   5640
         Width           =   615
      End
      Begin VB.Label lblUnit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   83
         Top             =   240
         Width           =   90
      End
      Begin VB.Label lblUnit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "2"
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   82
         Top             =   600
         Width           =   90
      End
      Begin VB.Label lblUnit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "3"
         Height          =   195
         Index           =   3
         Left            =   210
         TabIndex        =   81
         Top             =   960
         Width           =   90
      End
      Begin VB.Label lblUnit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "4"
         Height          =   195
         Index           =   4
         Left            =   210
         TabIndex        =   80
         Top             =   1320
         Width           =   90
      End
      Begin VB.Label lblUnit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "5"
         Height          =   195
         Index           =   5
         Left            =   210
         TabIndex        =   79
         Top             =   1680
         Width           =   90
      End
      Begin VB.Label lblUnit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "6"
         Height          =   195
         Index           =   6
         Left            =   210
         TabIndex        =   78
         Top             =   2040
         Width           =   90
      End
      Begin VB.Label lblUnit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "7"
         Height          =   195
         Index           =   7
         Left            =   210
         TabIndex        =   77
         Top             =   2400
         Width           =   90
      End
      Begin VB.Label lblUnit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "8"
         Height          =   195
         Index           =   8
         Left            =   210
         TabIndex        =   76
         Top             =   2760
         Width           =   90
      End
      Begin VB.Label lblUnit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "9"
         Height          =   195
         Index           =   9
         Left            =   210
         TabIndex        =   75
         Top             =   3120
         Width           =   90
      End
      Begin VB.Label lblUnit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "10"
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   74
         Top             =   3480
         Width           =   180
      End
      Begin VB.Label lblUnit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "11"
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   73
         Top             =   3840
         Width           =   180
      End
      Begin VB.Label lblUnit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "12"
         Height          =   195
         Index           =   12
         Left            =   120
         TabIndex        =   72
         Top             =   4200
         Width           =   180
      End
      Begin VB.Label lblUnit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "13"
         Height          =   195
         Index           =   13
         Left            =   120
         TabIndex        =   71
         Top             =   4560
         Width           =   180
      End
      Begin VB.Label lblUnit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "14"
         Height          =   195
         Index           =   14
         Left            =   120
         TabIndex        =   70
         Top             =   4920
         Width           =   180
      End
      Begin VB.Label lblUnit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "15"
         Height          =   195
         Index           =   15
         Left            =   120
         TabIndex        =   69
         Top             =   5280
         Width           =   180
      End
      Begin VB.Label lblUnit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "16"
         Height          =   195
         Index           =   16
         Left            =   120
         TabIndex        =   68
         Top             =   5640
         Width           =   180
      End
   End
   Begin VB.Timer timDim 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   3600
      Top             =   0
   End
   Begin VB.Timer timBright 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   2760
      Top             =   0
   End
   Begin VB.TextBox txtHouse 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   1
      Text            =   "D"
      Top             =   120
      Width           =   1095
   End
   Begin cm17a.controlcm ControlCm1 
      Height          =   30
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   53
   End
   Begin VB.Label lblCode 
      Caption         =   "House Code:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileMacros 
         Caption         =   "&Macros"
      End
      Begin VB.Menu mnuFileSchedule 
         Caption         =   "&Schedule"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTray 
      Caption         =   "TrayMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuTrayRestore 
         Caption         =   "Restore"
      End
      Begin VB.Menu mnuTrayExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmX10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBright_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    timBright.Enabled = False
End Sub

Private Sub cmdDim_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    timDim.Enabled = False
End Sub

Private Sub cmdOff_Click(Index As Integer)
    ControlCm1.Exec txtHouse, Str(Index), C_OFF
End Sub

Private Sub cmdDim_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        intUnit = Index
        timDim.Enabled = True
    End If
End Sub

Private Sub cmdOn_Click(Index As Integer)
    ControlCm1.Exec txtHouse, Str(Index), C_ON
End Sub

Private Sub cmdBright_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        intUnit = Index
        timBright.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    Dim init_error
    ControlCm1.comport = 1
    init_error = ControlCm1.Init
    If init_error <> 0 Then
        MsgBox "Error initializing CM17A", vbExclamation + vbOKOnly
    End If
    txtHouse.Text = GetSetting(App.Title, "Remote", "HouseCode", "A")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    i = X / Screen.TwipsPerPixelX
    Select Case i
        Case WM_LBUTTONDOWN Or WM_RBUTTONDOWN:
        Me.PopupMenu mnuTray
    End Select
    If i = WM_LBUTTONDBLCLK Then
        Me.WindowState = vbNormal
        Me.Show
        Me.Refresh
        Shell_NotifyIcon NIM_DELETE, nid
    End If
End Sub

Private Sub Form_Resize()
    If WindowState = vbMinimized Then
        Me.Hide
        Me.Refresh
        With nid
            .cbSize = Len(nid)
            .hwnd = Me.hwnd
            .uId = vbNull
            .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
            .uCallBackMessage = WM_MOUSEMOVE
            .hIcon = Me.Icon
            .szTip = Me.Caption & vbNullChar
        End With
        Shell_NotifyIcon NIM_ADD, nid
    Else
        Shell_NotifyIcon NIM_DELETE, nid
    End If
End Sub

Private Sub Form_Terminate()
    Shell_NotifyIcon NIM_DELETE, nid
    ControlCm1.ResetCom
    SaveSetting App.Title, "Remote", "HouseCode", txtHouse.Text
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.WindowState = vbMinimized
    Cancel = 1
End Sub

Private Sub mnuFileExit_Click()
    Call Form_Terminate
End Sub

Private Sub mnuFileMacros_Click()
    blnForce = True
    frmMacro.Show
    Me.Hide
End Sub

Private Sub mnuFileSchedule_Click()
    frmSchedule.Show
    frmSchedule.timFlash.Enabled = True
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuTrayExit_Click()
    Call Form_Terminate
End Sub

Private Sub mnuTrayRestore_Click()
    Me.WindowState = vbNormal
    Me.Show
    Me.Refresh
    Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub timBright_Timer()
    ControlCm1.Exec txtHouse, Str(intUnit), C_BRIGHT, 30
End Sub

Private Sub timDim_Timer()
    ControlCm1.Exec txtHouse, Str(intUnit), C_DIM, 30
End Sub

Private Sub txtHouse_Change()
    txtHouse.Text = UCase(txtHouse.Text)
    txtHouse.SelStart = 1
End Sub

Private Sub txtHouse_GotFocus()
    txtHouse.SelStart = 0
    txtHouse.SelLength = 1
End Sub
