Attribute VB_Name = "HeatMapReport"
Option Explicit

' =====================================================================
'  HEATMAP REPORT - controls ONLY the "HeatMap" sheet
'
'  RUN THIS:   HeatMap__Main    -> colours every % Cancelled cell compared with the month before
'                 GREEN = same or lower than the previous month (no increase)
'                 RED   = higher than the previous month (increase)
'                 first month of a therapist (nothing to compare) = no colour
'  IF NEEDED:  HeatMap__Check   (nothing happens - shows what the macro sees)
'
'  Runs by itself every time the HeatMap pivot changes (see ThisWorkbook)
'  Change anything in SETTINGS below - it affects the HeatMap sheet only
' =====================================================================

' ===== SETTINGS (change these) =====
Private Const H_SHEET As String = "HeatMap"             ' sheet name
Private Const H_TOLERANCE As Double = 0                 ' red only if the rise is bigger than this (0.02 = 2 points)
Private Const H_COLOUR_TOTAL_ROW As Boolean = True      ' also compare months in the Grand Total row
Private Const H_HEADER As Long = 16115137               ' header colour      = RGB(193,229,245) #C1E5F5
Private Const H_GREEN As Long = 13561798                ' no increase        = RGB(198,239,206)
Private Const H_GREEN_FONT As Long = 24832              ' green text         = RGB(0,97,0)
Private Const H_RED As Long = 13551615                  ' increase           = RGB(255,199,206)
Private Const H_RED_FONT As Long = 393372               ' red text           = RGB(156,0,6)
Private Const H_TOTAL As Long = 13431551                ' Grand Total column = RGB(255,242,204) #FFF2CC


' ===== MACROS TO RUN =====
Sub HeatMap__Main()
    Dim ws As Worksheet, pt As PivotTable
    Set ws = hSheet()
    If ws Is Nothing Then MsgBox "Sheet '" & H_SHEET & "' not found. Check the tab name.": Exit Sub
    Set pt = hPivot(ws)
    If pt Is Nothing Then MsgBox "No pivot table on sheet '" & H_SHEET & "'.": Exit Sub
    Application.ScreenUpdating = False
    hColour ws, pt
    Application.ScreenUpdating = True
End Sub

Sub HeatMap__Check()
    Dim ws As Worksheet, pt As PivotTable, hdr As Long, s As String
    Set ws = hSheet()
    If ws Is Nothing Then MsgBox "Sheet '" & H_SHEET & "' not found. Check the tab name.": Exit Sub
    Set pt = hPivot(ws)
    If pt Is Nothing Then MsgBox "No pivot table on sheet '" & H_SHEET & "'.": Exit Sub
    hdr = hHeaderRow(ws, pt)
    s = "Sheet: " & ws.Name & vbCrLf
    s = s & "Pivot: " & pt.Name & "  at " & pt.TableRange1.Address(False, False) & vbCrLf
    s = s & "Month header row: " & hdr & vbCrLf
    s = s & "First month: " & ws.Cells(hdr, pt.TableRange1.Column + 1).Text & vbCrLf
    s = s & "Last column header: " & ws.Cells(hdr, pt.TableRange1.Column + pt.TableRange1.Columns.Count - 1).Text
    MsgBox s
End Sub


' ===== HELPERS (not run by hand) =====
Private Function hSheet() As Worksheet
    On Error Resume Next
    Set hSheet = ThisWorkbook.Worksheets(H_SHEET)
    On Error GoTo 0
End Function

Private Function hPivot(ByVal ws As Worksheet) As PivotTable
    On Error Resume Next
    Set hPivot = ws.PivotTables(1)
    On Error GoTo 0
End Function

' the row with the month names (2024-11 | November ...)
Private Function hHeaderRow(ByVal ws As Worksheet, pt As PivotTable) As Long
    Dim r As Long, fr As Long, fc As Long
    fr = pt.TableRange1.Row: fc = pt.TableRange1.Column
    hHeaderRow = fr + 1
    For r = fr To fr + 3
        If ws.Cells(r, fc + 1).Text Like "#*" Then hHeaderRow = r: Exit Function
    Next r
End Function

Private Function hIsNumber(ByVal v As Variant) As Boolean
    If IsEmpty(v) Then Exit Function
    If IsError(v) Then Exit Function
    If VarType(v) = vbString Then Exit Function
    hIsNumber = IsNumeric(v)
End Function


' ===== COLOURS =====
Private Sub hColour(ByVal ws As Worksheet, pt As PivotTable)
    Dim fr As Long, lr As Long, fc As Long, lc As Long
    Dim hdr As Long, d1 As Long, dLast As Long, totalRow As Long
    Dim r As Long, c As Long, v As Variant, prev As Double, hasPrev As Boolean
    Dim cell As Range

    fr = pt.TableRange1.Row: fc = pt.TableRange1.Column
    lr = fr + pt.TableRange1.Rows.Count - 1
    lc = fc + pt.TableRange1.Columns.Count - 1
    hdr = hHeaderRow(ws, pt)
    d1 = hdr + 1

    ' Grand Total column / row (if the pivot shows them)
    dLast = lc
    If InStr(1, ws.Cells(hdr, lc).Text, "Grand Total", vbTextCompare) > 0 Then dLast = lc - 1
    If InStr(1, ws.Cells(lr, fc).Text, "Grand Total", vbTextCompare) > 0 Then totalRow = lr

    ' clean slate
    With ws.Range(ws.Cells(fr, fc), ws.Cells(lr, lc))
        .Interior.ColorIndex = xlNone
        .Font.ColorIndex = xlAutomatic
        .Font.Bold = False
    End With

    ' header rows
    With ws.Range(ws.Cells(fr, fc), ws.Cells(hdr, lc))
        .Interior.Color = H_HEADER
        .Font.Bold = True
    End With

    ' each therapist: compare every month with the previous month that has a value
    For r = d1 To lr
        If r = totalRow And Not H_COLOUR_TOTAL_ROW Then GoTo nextRow
        hasPrev = False
        For c = fc + 1 To dLast
            Set cell = ws.Cells(r, c)
            v = cell.Value
            If hIsNumber(v) Then
                If hasPrev Then
                    If Round(CDbl(v), 4) > Round(prev, 4) + H_TOLERANCE Then
                        cell.Interior.Color = H_RED
                        cell.Font.Color = H_RED_FONT
                    Else
                        cell.Interior.Color = H_GREEN
                        cell.Font.Color = H_GREEN_FONT
                    End If
                End If
                prev = CDbl(v)
                hasPrev = True
            End If
        Next c
nextRow:
    Next r

    ' Grand Total row and column: bold, light gold where not coloured
    If totalRow > 0 Then
        ws.Range(ws.Cells(totalRow, fc), ws.Cells(totalRow, lc)).Font.Bold = True
        ws.Cells(totalRow, fc).Interior.Color = H_TOTAL
        If Not H_COLOUR_TOTAL_ROW Then ws.Range(ws.Cells(totalRow, fc), ws.Cells(totalRow, lc)).Interior.Color = H_TOTAL
    End If
    If dLast < lc Then
        With ws.Range(ws.Cells(d1, lc), ws.Cells(lr, lc))
            .Interior.Color = H_TOTAL
            .Font.Bold = True
        End With
    End If

    ws.Range(ws.Cells(fr, fc), ws.Cells(lr, lc)).Columns.AutoFit
End Sub
