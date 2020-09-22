VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmSchedule 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scheduling Options"
   ClientHeight    =   6240
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   8850
   Icon            =   "frmSchedule.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   8850
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "Clear &All Timeslots"
      Height          =   375
      Left            =   7080
      TabIndex        =   53
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear Timeslot"
      Height          =   375
      Left            =   5160
      TabIndex        =   52
      Top             =   720
      Width           =   1815
   End
   Begin VB.Timer timFlash 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2760
      Top             =   3240
   End
   Begin VB.Timer timCheck 
      Interval        =   3000
      Left            =   3360
      Top             =   3240
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   3960
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "&Change Marco File"
      Height          =   375
      Left            =   2760
      TabIndex        =   50
      Top             =   720
      Width           =   2295
   End
   Begin VB.Frame fraTime 
      Caption         =   "Scheduleing Times:"
      Height          =   6135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2655
      Begin VB.OptionButton optTime 
         Caption         =   "11:30"
         Height          =   255
         Index           =   47
         Left            =   1320
         TabIndex        =   48
         Top             =   5760
         Width           =   1215
      End
      Begin VB.OptionButton optTime 
         Caption         =   "11:00"
         Height          =   255
         Index           =   46
         Left            =   1320
         TabIndex        =   47
         Top             =   5520
         Width           =   1215
      End
      Begin VB.OptionButton optTime 
         Caption         =   "10:30"
         Height          =   255
         Index           =   45
         Left            =   1320
         TabIndex        =   46
         Top             =   5280
         Width           =   1215
      End
      Begin VB.OptionButton optTime 
         Caption         =   "10:00"
         Height          =   255
         Index           =   44
         Left            =   1320
         TabIndex        =   45
         Top             =   5040
         Width           =   1215
      End
      Begin VB.OptionButton optTime 
         Caption         =   "9:30"
         Height          =   255
         Index           =   43
         Left            =   1320
         TabIndex        =   44
         Top             =   4800
         Width           =   1215
      End
      Begin VB.OptionButton optTime 
         Caption         =   "9:00"
         Height          =   255
         Index           =   42
         Left            =   1320
         TabIndex        =   43
         Top             =   4560
         Width           =   1215
      End
      Begin VB.OptionButton optTime 
         Caption         =   "8:30"
         Height          =   255
         Index           =   41
         Left            =   1320
         TabIndex        =   42
         Top             =   4320
         Width           =   1215
      End
      Begin VB.OptionButton optTime 
         Caption         =   "8:00"
         Height          =   255
         Index           =   40
         Left            =   1320
         TabIndex        =   41
         Top             =   4080
         Width           =   1215
      End
      Begin VB.OptionButton optTime 
         Caption         =   "7:30"
         Height          =   255
         Index           =   39
         Left            =   1320
         TabIndex        =   40
         Top             =   3840
         Width           =   1215
      End
      Begin VB.OptionButton optTime 
         Caption         =   "7:00"
         Height          =   255
         Index           =   38
         Left            =   1320
         TabIndex        =   39
         Top             =   3600
         Width           =   1215
      End
      Begin VB.OptionButton optTime 
         Caption         =   "6:30"
         Height          =   255
         Index           =   37
         Left            =   1320
         TabIndex        =   38
         Top             =   3360
         Width           =   1215
      End
      Begin VB.OptionButton optTime 
         Caption         =   "6:00"
         Height          =   255
         Index           =   36
         Left            =   1320
         TabIndex        =   37
         Top             =   3120
         Width           =   1215
      End
      Begin VB.OptionButton optTime 
         Caption         =   "5:30"
         Height          =   255
         Index           =   35
         Left            =   1320
         TabIndex        =   36
         Top             =   2880
         Width           =   1215
      End
      Begin VB.OptionButton optTime 
         Caption         =   "5:00"
         Height          =   255
         Index           =   34
         Left            =   1320
         TabIndex        =   35
         Top             =   2640
         Width           =   1215
      End
      Begin VB.OptionButton optTime 
         Caption         =   "4:30"
         Height          =   255
         Index           =   33
         Left            =   1320
         TabIndex        =   34
         Top             =   2400
         Width           =   1215
      End
      Begin VB.OptionButton optTime 
         Caption         =   "4:00"
         Height          =   255
         Index           =   32
         Left            =   1320
         TabIndex        =   33
         Top             =   2160
         Width           =   1215
      End
      Begin VB.OptionButton optTime 
         Caption         =   "3:30"
         Height          =   255
         Index           =   31
         Left            =   1320
         TabIndex        =   32
         Top             =   1920
         Width           =   1215
      End
      Begin VB.OptionButton optTime 
         Caption         =   "3:00"
         Height          =   255
         Index           =   30
         Left            =   1320
         TabIndex        =   31
         Top             =   1680
         Width           =   1215
      End
      Begin VB.OptionButton optTime 
         Caption         =   "2:30"
         Height          =   255
         Index           =   29
         Left            =   1320
         TabIndex        =   30
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton optTime 
         Caption         =   "2:00"
         Height          =   255
         Index           =   28
         Left            =   1320
         TabIndex        =   29
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton optTime 
         Caption         =   "1:30"
         Height          =   255
         Index           =   27
         Left            =   1320
         TabIndex        =   28
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton optTime 
         Caption         =   "1:00"
         Height          =   255
         Index           =   26
         Left            =   1320
         TabIndex        =   27
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optTime 
         Caption         =   "12:30"
         Height          =   255
         Index           =   25
         Left            =   1320
         TabIndex        =   26
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton optTime 
         Caption         =   "Noon"
         Height          =   255
         Index           =   24
         Left            =   1320
         TabIndex        =   25
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optTime 
         Caption         =   "11:30"
         Height          =   255
         Index           =   23
         Left            =   120
         TabIndex        =   24
         Top             =   5760
         Width           =   1215
      End
      Begin VB.OptionButton optTime 
         Caption         =   "11:00"
         Height          =   255
         Index           =   22
         Left            =   120
         TabIndex        =   23
         Top             =   5520
         Width           =   1215
      End
      Begin VB.OptionButton optTime 
         Caption         =   "10:30"
         Height          =   255
         Index           =   21
         Left            =   120
         TabIndex        =   22
         Top             =   5280
         Width           =   1215
      End
      Begin VB.OptionButton optTime 
         Caption         =   "10:00"
         Height          =   255
         Index           =   20
         Left            =   120
         TabIndex        =   21
         Top             =   5040
         Width           =   1215
      End
      Begin VB.OptionButton optTime 
         Caption         =   "9:30"
         Height          =   255
         Index           =   19
         Left            =   120
         TabIndex        =   20
         Top             =   4800
         Width           =   1215
      End
      Begin VB.OptionButton optTime 
         Caption         =   "9:00"
         Height          =   255
         Index           =   18
         Left            =   120
         TabIndex        =   19
         Top             =   4560
         Width           =   1215
      End
      Begin VB.OptionButton optTime 
         Caption         =   "8:30"
         Height          =   255
         Index           =   17
         Left            =   120
         TabIndex        =   18
         Top             =   4320
         Width           =   1215
      End
      Begin VB.OptionButton optTime 
         Caption         =   "8:00"
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   17
         Top             =   4080
         Width           =   1215
      End
      Begin VB.OptionButton optTime 
         Caption         =   "7:30"
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   16
         Top             =   3840
         Width           =   1215
      End
      Begin VB.OptionButton optTime 
         Caption         =   "7:00"
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   15
         Top             =   3600
         Width           =   1215
      End
      Begin VB.OptionButton optTime 
         Caption         =   "6:30"
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   14
         Top             =   3360
         Width           =   1215
      End
      Begin VB.OptionButton optTime 
         Caption         =   "6:00"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   13
         Top             =   3120
         Width           =   1215
      End
      Begin VB.OptionButton optTime 
         Caption         =   "5:30"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   12
         Top             =   2880
         Width           =   1215
      End
      Begin VB.OptionButton optTime 
         Caption         =   "5:00"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   11
         Top             =   2640
         Width           =   1215
      End
      Begin VB.OptionButton optTime 
         Caption         =   "4:30"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   10
         Top             =   2400
         Width           =   1215
      End
      Begin VB.OptionButton optTime 
         Caption         =   "4:00"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   9
         Top             =   2160
         Width           =   1215
      End
      Begin VB.OptionButton optTime 
         Caption         =   "3:30"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   1215
      End
      Begin VB.OptionButton optTime 
         Caption         =   "3:00"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   1215
      End
      Begin VB.OptionButton optTime 
         Caption         =   "2:30"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton optTime 
         Caption         =   "2:00"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton optTime 
         Caption         =   "1:30"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton optTime 
         Caption         =   "1:00"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optTime 
         Caption         =   "12:30"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton optTime 
         Caption         =   "Midnight"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Label lblWarning 
      Caption         =   $"frmSchedule.frx":0442
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1695
      Left            =   2760
      TabIndex        =   51
      Top             =   1320
      Width           =   5655
   End
   Begin VB.Label lblFile 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2760
      TabIndex        =   49
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "frmSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intSelected As Integer
Dim strHouse As String, strUnit As String, strCommand As String

Private Sub cmdChange_Click()
    'changes or adds a macro to br ran at the selected time block
    On Error GoTo Error
    Dialog.CancelError = True
    Dialog.Flags = cdlOFNFileMustExist
    Dialog.Filter = "X10 Macro Files (*.xmf) | *.xmf"
    Dialog.ShowOpen
    udtSchedule.strFile = Dialog.FileName
    Open App.Path & "\ScheduleData.dat" For Random As #1 Len = Len(udtSchedule)
    Put #1, intSelected + 1, udtSchedule
    Close #1
    lblFile.Caption = Dialog.FileName
    Call MemSave
Exit Sub
Error:
End Sub

Private Sub cmdClear_Click()
    'clears selected time block
    udtSchedule.strFile = ""
    Open App.Path & "\ScheduleData.dat" For Random As #1 Len = Len(udtSchedule)
    Put #1, intSelected + 1, udtSchedule
    Close #1
    lblFile.Caption = ""
    Call MemSave
End Sub

Private Sub cmdClearAll_Click()
    'clears all scheduled macros
    Dim intMsg As Integer
    intMsg = MsgBox("Are you sure you want to delete all macro entries?", vbQuestion + vbYesNo + vbApplicationModal, "Idoit Proof")
    If intMsg = vbNo Then Exit Sub
    udtSchedule.strFile = ""
    Open App.Path & "\ScheduleData.dat" For Random As #1 Len = Len(udtSchedule)
    For i = 1 To 48 Step 1
        Put #1, i, udtSchedule
    Next i
    Close #1
    lblFile.Caption = ""
    Call MemSave
End Sub

Private Sub Form_Load()
    'prevents program from accessing hard drive every half hour
    Call MemSave
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'don't let user unload form because then it won't run any scheduled macros
    Me.Hide
    timCheck.Enabled = True
    Cancel = 1
End Sub

Private Sub optTime_Click(Index As Integer)
    'opens the marco (if any) that is scheduled for the selected time
    intSelected = Index
    Open App.Path & "\ScheduleData.dat" For Random As #1 Len = Len(udtSchedule)
    Get #1, Index + 1, udtSchedule
    Close #1
    lblFile.Caption = udtSchedule.strFile
End Sub

Private Sub timCheck_Timer()
    'this is how the program finds out what time it is
    'the integer represents the aray numbers of the option buttons
    Select Case Left(Time, 5)
        Case Is = "1:00:"
            If Right(Time, 2) = "AM" Then Call Execute(2) Else: Call Execute(26)
        Case Is = "1:30:"
            If Right(Time, 2) = "AM" Then Call Execute(3) Else: Call Execute(27)
        Case Is = "2:00:"
            If Right(Time, 2) = "AM" Then Call Execute(4) Else: Call Execute(28)
        Case Is = "2:30:"
            If Right(Time, 2) = "AM" Then Call Execute(5) Else: Call Execute(29)
        Case Is = "3:00:"
            If Right(Time, 2) = "AM" Then Call Execute(6) Else: Call Execute(30)
        Case Is = "3:30:"
            If Right(Time, 2) = "AM" Then Call Execute(7) Else: Call Execute(31)
        Case Is = "4:00:"
            If Right(Time, 2) = "AM" Then Call Execute(8) Else: Call Execute(32)
        Case Is = "4:30:"
            If Right(Time, 2) = "AM" Then Call Execute(9) Else: Call Execute(33)
        Case Is = "5:30:"
            If Right(Time, 2) = "AM" Then Call Execute(10) Else: Call Execute(34)
        Case Is = "5:00:"
            If Right(Time, 2) = "AM" Then Call Execute(11) Else: Call Execute(35)
        Case Is = "6:00:"
            If Right(Time, 2) = "AM" Then Call Execute(12) Else: Call Execute(36)
        Case Is = "6:30:"
            If Right(Time, 2) = "AM" Then Call Execute(13) Else: Call Execute(37)
        Case Is = "7:00:"
            If Right(Time, 2) = "AM" Then Call Execute(14) Else: Call Execute(38)
        Case Is = "7:30:"
            If Right(Time, 2) = "AM" Then Call Execute(15) Else: Call Execute(39)
        Case Is = "8:00:"
            If Right(Time, 2) = "AM" Then Call Execute(16) Else: Call Execute(40)
        Case Is = "8:30:"
            If Right(Time, 2) = "AM" Then Call Execute(17) Else: Call Execute(41)
        Case Is = "9:00:"
            If Right(Time, 2) = "AM" Then Call Execute(18) Else: Call Execute(42)
        Case Is = "9:30:"
            If Right(Time, 2) = "AM" Then Call Execute(19) Else: Call Execute(43)
        Case Is = "10:00"
            If Right(Time, 2) = "AM" Then Call Execute(20) Else: Call Execute(44)
        Case Is = "10:30"
            If Right(Time, 2) = "AM" Then Call Execute(21) Else: Call Execute(45)
        Case Is = "11:00"
            If Right(Time, 2) = "AM" Then Call Execute(22) Else: Call Execute(46)
        Case Is = "11:30"
            If Right(Time, 2) = "AM" Then Call Execute(23) Else: Call Execute(47)
        Case Is = "12:00"
            If Right(Time, 2) = "AM" Then Call Execute(0) Else: Call Execute(24)
        Case Is = "12:30"
            If Right(Time, 2) = "AM" Then Call Execute(1) Else: Call Execute(25)
    End Select
End Sub

Private Sub Execute(intTime As Integer)
    On Error GoTo SError
    'check to see if there is a scheduled marco for this time
    If strMacros(intTime + 1) = "" Then
        Exit Sub
        timCheck.Interval = 3000
    End If
    'make sure the marco exists
    If Dir(strMacros(intTime + 1)) = "" Then Exit Sub
    'if it made it this far, then there is a scheduled macro for this time
    blnSchedule = True
    timCheck.Interval = 60000
    strFileName = strMacros(intTime + 1)
    Dim frmSMacro As New frmMacro
    Load frmSMacro
    frmSMacro.Left = 0
    frmSMacro.Top = 0
    frmSMacro.Caption = "Scheduled Macro in Progress..."
    frmSMacro.Show
    With frmSMacro
        Open strFileName For Input As #1
        .lstCommands.Clear
        Do While Not EOF(1)
            Input #1, strHouse, strUnit, strCommand
            .lstCommands.AddItem strHouse & ", " & strUnit & ", " & strCommand
        Loop
        Close #1
    End With
    Call frmSMacro.cmdRun_Click
Exit Sub
SError:
    MsgBox "Error running scheduled macro.  (" & Err.Description & ", " & strFileName & ")", vbCritical, "Macro Error"
End Sub

'This function saves the schedule file into variables
'so it doesn't keep the hard drive spinning
Private Sub MemSave()
    Open App.Path & "\ScheduleData.dat" For Random As #1 Len = Len(udtSchedule)
    For i = 1 To 48
        Get #1, i, udtSchedule
        strMacros(i) = Trim(udtSchedule.strFile)
    Next i
    Close #1
End Sub

Private Sub timFlash_Timer()
    Static intFlash As Integer
    If lblWarning.Visible = True Then lblWarning.Visible = False Else:  lblWarning.Visible = True
    intFlash = intFlash + 1
    If intFlash = 6 Then
        timFlash.Enabled = False
        intFlash = 0
    End If
End Sub
