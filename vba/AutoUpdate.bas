Attribute VB_Name = "AutoUpdate"
Option Explicit

' =====================================================================
'  AUTO UPDATE - makes everything run by itself
'
'  Auto_Open      -> runs BY ITSELF when the file opens:
'                    switches automatic updates on, shows latest week + month, colours heat map
'  Auto__Start    -> run by hand (Alt+F8) if things stop updating
'                    (needed after you edit VBA code or after a macro error)
'  Auto__Stop     -> switches automatic updates off
'  Auto__Status   -> shows if automatic updates are on
'
'  Needs the CLASS module "AutoEvents"
' =====================================================================

Private mAuto As AutoEvents

' Excel runs a macro called Auto_Open by itself when the file is opened
Sub Auto_Open()
    Auto__Start
    aRun "Weekly__Main"            ' Weekly 2026 sheet
    aRun "Monthly__Main"           ' Monthly 2026 sheet
    aRun "HeatMap__Main"           ' HeatMap sheet
End Sub

Sub Auto__Start()
    Application.EnableEvents = True
    Set mAuto = New AutoEvents
    Set mAuto.Wb = ThisWorkbook
End Sub

Sub Auto__Stop()
    Set mAuto = Nothing
    MsgBox "Automatic updates are OFF." & vbCrLf & "Run Auto__Start to switch them on again."
End Sub

Sub Auto__Status()
    If mAuto Is Nothing Or Not Application.EnableEvents Then
        MsgBox "Automatic updates are OFF." & vbCrLf & "Run Auto__Start to switch them on."
    Else
        MsgBox "Automatic updates are ON."
    End If
End Sub


' ===== HELPER (not run by hand) =====
Private Sub aRun(ByVal macroName As String)
    On Error Resume Next
    Application.Run "'" & ThisWorkbook.Name & "'!" & macroName
    On Error GoTo 0
End Sub
