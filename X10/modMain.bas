Attribute VB_Name = "modMain"
Public intUnit As Integer, strCode As String, strFileName As String, strScheduleName As String
Public blnLoop As Boolean, blnCount As Boolean, intTime As Integer, strTime As String, i As Integer
Public blnForce As Boolean, blnSchedule As Boolean
Public strMacros(1 To 48) As String

Type Schedule
    strFile As String * 150
End Type
Public udtSchedule As Schedule

Sub Main()
    blnForce = False
    Load frmMacro
    Load frmSchedule
    frmSchedule.timCheck.Enabled = True
    Load frmX10
    frmX10.Show
End Sub
