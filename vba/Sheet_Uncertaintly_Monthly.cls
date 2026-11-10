' SHEET MODULE of "Uncertaintly_Monthly" (not a normal module):
' Alt+F11 > Microsoft Excel Objects > double-click the sheet > paste the code below.

Option Explicit

' Recolours this sheet and updates the month box every time the slicer is clicked
Private Sub Worksheet_PivotTableUpdate(ByVal Target As PivotTable)
    Application.EnableEvents = False
    Application.Run "'" & ThisWorkbook.Name & "'!UncertMonthly__Recolour"
    Application.EnableEvents = True
End Sub
