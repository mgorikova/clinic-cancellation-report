' SHEET MODULE of "Uncertaintly_Weekly" (not a normal module):
' Alt+F11 > Microsoft Excel Objects > double-click the sheet > paste the code below.

Option Explicit

' Recolours this sheet and updates the week box every time the slicer is clicked
Private Sub Worksheet_PivotTableUpdate(ByVal Target As PivotTable)
    Application.EnableEvents = False
    Application.Run "'" & ThisWorkbook.Name & "'!UncertWeekly__Recolour"
    Application.EnableEvents = True
End Sub
