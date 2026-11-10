VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AutoEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =====================================================================
'  AutoEvents  (CLASS module - Insert > Class Module, name it AutoEvents)
'  Listens to the workbook: every time a pivot changes (buttons OR
'  clicking a slicer yourself) it recolours that sheet and updates the box.
'  Switched on by the AUTO_UPDATE module - nothing to run here.
'  Macros are called by NAME, so a missing one is simply skipped (no compile error).
' =====================================================================

Public WithEvents Wb As Workbook

Private Sub Wb_SheetPivotTableUpdate(ByVal sh As Object, ByVal Target As PivotTable)
    Application.EnableEvents = False          ' stop it running twice
    Select Case sh.Name
        Case "Weekly 2026":  RunIfThere "Weekly__Recolour"
        Case "Monthly 2026": RunIfThere "Monthly__Recolour"
        Case "HeatMap":      RunIfThere "HeatMap__Main"
    End Select
    Application.EnableEvents = True           ' always switch automatic updates back on
End Sub

Private Sub RunIfThere(ByVal macroName As String)
    On Error Resume Next
    Application.Run "'" & ThisWorkbook.Name & "'!" & macroName
    On Error GoTo 0
End Sub
