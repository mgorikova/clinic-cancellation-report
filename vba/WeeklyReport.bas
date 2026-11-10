Attribute VB_Name = "WeeklyReport"
Option Explicit

' =====================================================================
'  WEEKLY REPORT - controls ONLY the "Weekly 2026" sheet
'
'  RUN THIS:   Weekly__Main        -> copies the Monthly look, jumps to the latest week,
'                                     colours the table, draws the week box
'  BUTTONS:    <  = Weekly__Previous      >  = Weekly__Next
'  IF NEEDED:  Weekly__Recolour   (colours look wrong)
'              Weekly__FixButtons (a button got squashed)
'              Weekly__Check      (nothing happens - shows what the macro sees)
'
'  Change anything in SETTINGS below - it affects the Weekly 2026 sheet only
' =====================================================================

' ===== SETTINGS (change these) =====
Private Const W_SHEET As String = "Weekly 2026"         ' sheet name
Private Const W_SLICER As String = "Slicer_Year_Month"  ' slicer (if not found, uses the slicer on the sheet)
Private Const W_YEAR As String = "2026"                 ' only weeks of this year ("" = all years)
Private Const W_BOX_CELL As String = "J4"               ' week box always sits here
Private Const W_BOX_NAME As String = "lblPeriod"        ' name of the week box shape
Private Const W_BOX_WIDTH As String = "slicer"          ' week box width: "slicer" or "button"
Private Const W_SLICER_CELL As String = "J6"            ' slicer top-left corner stays here
Private Const W_STYLE_FROM As String = "Monthly 2026"   ' copy button look/size + top row heights from this sheet ("" = off)
Private Const W_FONT_SIZE As Long = 20                  ' max text size in the week box
Private Const W_TEXT_COLOR As Long = 0                  ' week box text colour (0 = black, 16777215 = white)
Private Const W_THRESHOLD As Double = 0.2               ' orange row if % Cancelled > 20%
Private Const W_HEADER As Long = 16115137               ' header colour        = RGB(193,229,245) #C1E5F5
Private Const W_ORANGE As Long = 11851260               ' high cancel row      = RGB(252,213,180)
Private Const W_TOTAL As Long = 13431551                ' Grand Total row      = RGB(255,242,204) #FFF2CC
Private Const W_LINE As Long = 6567967                  ' box outline (navy)   = RGB(31,56,100)


' ===== MACROS TO RUN =====
Sub Weekly__Main()
    wCopyStyle
    wMove 0
End Sub

Sub Weekly__Previous()
    wMove -1
End Sub

Sub Weekly__Next()
    wMove 1
End Sub

Sub Weekly__Recolour()
    Dim ws As Worksheet, sc As SlicerCache
    Set ws = wSheet()
    If ws Is Nothing Then Exit Sub
    wCopyStyle
    wFormat ws
    Set sc = wSlicer(ws)
    If Not sc Is Nothing Then wShowBox ws, sc
End Sub

Sub Weekly__FixButtons()
    Dim ws As Worksheet
    Set ws = wSheet()
    If ws Is Nothing Then Exit Sub
    wKeepButtonsSafe ws
    Weekly__Recolour
End Sub

Sub Weekly__Check()
    Dim ws As Worksheet, sc As SlicerCache, s As String, i As Long
    Set ws = wSheet()
    If ws Is Nothing Then MsgBox "Sheet '" & W_SHEET & "' not found. Check the tab name.": Exit Sub
    s = "Sheet: " & ws.Name & vbCrLf & "Pivots on sheet: " & ws.PivotTables.Count & vbCrLf
    Set sc = wSlicer(ws)
    If sc Is Nothing Then MsgBox s & "NO slicer on this sheet.": Exit Sub
    s = s & "Slicer used: " & sc.Name & "  (" & sc.SourceName & ")" & vbCrLf
    s = s & "First item: " & wFirstCaption(sc) & vbCrLf
    s = s & "Items: " & sc.SlicerCacheLevels(1).SlicerItems.Count & vbCrLf
    s = s & "Selected index: " & wSelected(sc) & vbCrLf
    i = wLatest(sc)
    s = s & "Latest index: " & i
    If i > 0 Then s = s & " = " & sc.SlicerCacheLevels(1).SlicerItems(i).Caption
    s = s & vbCrLf & "< button found: " & CStr(Not wButton(ws, False) Is Nothing)
    s = s & vbCrLf & "> button found: " & CStr(Not wButton(ws, True) Is Nothing)
    MsgBox s
End Sub


' ===== SAME STYLE AS MONTHLY (helpers - not run by hand) =====
' makes the Weekly < > buttons look exactly like the Monthly ones (shape, colour, size, gap)
' and gives the top rows (above the slicer) the same heights as on the Monthly sheet
Private Sub wCopyStyle()
    Dim ws As Worksheet, src As Worksheet
    Dim sp As Object, sn As Object, bp As Object, bn As Object
    Dim r As Long, lastRow As Long, gap As Double

    If W_STYLE_FROM = "" Then Exit Sub
    Set ws = wSheet()
    On Error Resume Next
    Set src = ThisWorkbook.Worksheets(W_STYLE_FROM)
    On Error GoTo 0
    If ws Is Nothing Or src Is Nothing Then Exit Sub

    Set sp = wButton(src, False): Set sn = wButton(src, True)
    Set bp = wButton(ws, False):  Set bn = wButton(ws, True)
    If sp Is Nothing Or sn Is Nothing Or bp Is Nothing Or bn Is Nothing Then Exit Sub

    ' 1) same row heights above the slicer (rows 1-5)
    lastRow = ws.Range(W_SLICER_CELL).Row - 1
    For r = 1 To lastRow
        ws.Rows(r).RowHeight = src.Rows(r).RowHeight
    Next r

    ' 2) same shape + look (fill, outline, text style) as the Monthly buttons
    On Error Resume Next
    bp.AutoShapeType = sp.AutoShapeType
    bn.AutoShapeType = sn.AutoShapeType
    sp.PickUp: bp.Apply
    sn.PickUp: bn.Apply
    bp.TextFrame2.TextRange.Font.Size = sp.TextFrame2.TextRange.Font.Size
    bn.TextFrame2.TextRange.Font.Size = sn.TextFrame2.TextRange.Font.Size
    bp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = sp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
    bn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = sn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
    On Error GoTo 0

    ' 3) same size, same height on the sheet, same gap between < and >
    bp.LockAspectRatio = msoFalse: bn.LockAspectRatio = msoFalse
    bp.Placement = xlFreeFloating: bn.Placement = xlFreeFloating
    bp.Width = sp.Width:  bp.Height = sp.Height
    bn.Width = sn.Width:  bn.Height = sn.Height
    bp.Top = sp.Top:      bn.Top = sn.Top
    gap = sn.Left - (sp.Left + sp.Width)
    bn.Left = bp.Left + bp.Width + gap
End Sub


' ===== NAVIGATION (helpers - not run by hand) =====
Private Function wSheet() As Worksheet
    On Error Resume Next
    Set wSheet = ThisWorkbook.Worksheets(W_SHEET)
    On Error GoTo 0
End Function

Private Function wSlicer(ByVal ws As Worksheet) As SlicerCache
    ' 1st: slicer on this sheet with items like "2026-01-05"
    ' 2nd: slicer named in SETTINGS (if on this sheet)   3rd: any slicer on this sheet
    Dim sc As SlicerCache, firstOnSheet As SlicerCache, named As SlicerCache
    For Each sc In ThisWorkbook.SlicerCaches
        If wOnSheet(sc, ws) Then
            If firstOnSheet Is Nothing Then Set firstOnSheet = sc
            If sc.Name = W_SLICER Then Set named = sc
            If wFirstCaption(sc) Like "####-##-##" Then Set wSlicer = sc: Exit Function
        End If
    Next sc
    If Not named Is Nothing Then Set wSlicer = named: Exit Function
    Set wSlicer = firstOnSheet
End Function

Private Function wOnSheet(sc As SlicerCache, ByVal ws As Worksheet) As Boolean
    Dim i As Long
    On Error Resume Next
    For i = 1 To sc.Slicers.Count
        If sc.Slicers(i).Shape.Parent.Name = ws.Name Then wOnSheet = True: Exit Function
    Next i
End Function

Private Function wFirstCaption(sc As SlicerCache) As String
    Dim itms As SlicerItems, i As Long
    On Error Resume Next
    Set itms = sc.SlicerCacheLevels(1).SlicerItems
    If itms Is Nothing Then Exit Function
    For i = 1 To itms.Count
        If Trim(itms(i).Caption) <> "" Then wFirstCaption = itms(i).Caption: Exit Function
    Next i
End Function

Private Function wInScope(it As SlicerItem) As Boolean
    If Not it.HasData Then Exit Function
    If Not (it.Caption Like "#*") Then Exit Function      ' skip "(blank)" and other non-date items
    If W_YEAR <> "" And it.Caption Like "####*" Then      ' year filter only if caption starts with a year
        If Left(it.Caption, 4) <> W_YEAR Then Exit Function
    End If
    wInScope = True
End Function

Private Function wSelected(sc As SlicerCache) As Long
    Dim sel As Variant, cur As String, i As Long, itms As SlicerItems
    Set itms = sc.SlicerCacheLevels(1).SlicerItems
    On Error Resume Next
    sel = sc.VisibleSlicerItemsList
    On Error GoTo 0
    If IsArray(sel) Then
        If UBound(sel) = LBound(sel) Then cur = CStr(sel(LBound(sel)))
    End If
    If cur = "" Then Exit Function
    For i = 1 To itms.Count
        If itms(i).Name = cur Then wSelected = i: Exit Function
    Next i
End Function

Private Function wStep(sc As SlicerCache, ByVal idx As Long, ByVal stepBy As Long) As Long
    ' Previous = closest EARLIER week, Next = closest LATER week.
    ' Compares the week names themselves, so it works whatever order the slicer shows them in.
    Dim itms As SlicerItems, i As Long, cur As String, cap As String, best As String
    Set itms = sc.SlicerCacheLevels(1).SlicerItems
    cur = itms(idx).Caption
    For i = 1 To itms.Count
        If wInScope(itms(i)) Then
            cap = itms(i).Caption
            If stepBy < 0 Then
                If cap < cur Then
                    If wStep = 0 Or cap > best Then wStep = i: best = cap
                End If
            Else
                If cap > cur Then
                    If wStep = 0 Or cap < best Then wStep = i: best = cap
                End If
            End If
        End If
    Next i
End Function

Private Function wLatest(sc As SlicerCache) As Long
    ' the newest week that has data (and is in the year from SETTINGS)
    Dim itms As SlicerItems, i As Long, best As String
    Set itms = sc.SlicerCacheLevels(1).SlicerItems
    For i = 1 To itms.Count
        If wInScope(itms(i)) Then
            If wLatest = 0 Or itms(i).Caption > best Then wLatest = i: best = itms(i).Caption
        End If
    Next i
End Function

Private Sub wMove(ByVal stepBy As Long)
    Dim ws As Worksheet, sc As SlicerCache, idx As Long, i As Long
    Set ws = wSheet()
    If ws Is Nothing Then MsgBox "Sheet '" & W_SHEET & "' not found.": Exit Sub
    Set sc = wSlicer(ws)
    If sc Is Nothing Then MsgBox "No slicer found on sheet '" & W_SHEET & "'.": Exit Sub

    idx = wSelected(sc)
    If idx > 0 Then
        If Not wInScope(sc.SlicerCacheLevels(1).SlicerItems(idx)) Then idx = 0
    End If
    If stepBy = 0 Or idx = 0 Then
        i = wLatest(sc)
    Else
        i = wStep(sc, idx, stepBy)
    End If
    If i = 0 Then
        If stepBy = 0 Or idx = 0 Then
            MsgBox "No week found to show." & vbCrLf & _
                   "Slicer used: " & sc.Name & vbCrLf & _
                   "First item: " & wFirstCaption(sc) & vbCrLf & vbCrLf & _
                   "Run Weekly__Check for details."
        Else
            Beep                                            ' already at first / last week
        End If
        Exit Sub
    End If
    wSelect ws, sc, i
End Sub

Private Sub wSelect(ByVal ws As Worksheet, sc As SlicerCache, ByVal i As Long)
    Dim itemName As String, cap As String, ok As Boolean, errText As String
    itemName = sc.SlicerCacheLevels(1).SlicerItems(i).Name
    cap = sc.SlicerCacheLevels(1).SlicerItems(i).Caption

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error Resume Next
    sc.VisibleSlicerItemsList = Array(itemName)
    ok = (Err.Number = 0)
    errText = Err.Description
    Err.Clear
    If Not ok Then
        ws.PivotTables(1).CubeFields(sc.SourceName).PivotFields(1).VisibleItemsList = Array(itemName)
        ok = (Err.Number = 0)
        Err.Clear
    End If
    On Error GoTo 0
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If Not ok Then MsgBox "Could not select " & cap & "." & vbCrLf & errText: Exit Sub
    wFormat ws
    wShowBox ws, sc
End Sub


' ===== WEEK BOX (fixed at J4, looks like the < > buttons, width = slicer; slicer fixed at J6) =====
Private Function wLabel(ByVal cap As String) As String
    ' "2026-01-05" -> "2026-01-05  (5 Jan - 9 Jan 2026)"
    Dim d As Date
    If cap Like "####-##-##" Then
        d = DateSerial(CLng(Left(cap, 4)), CLng(Mid(cap, 6, 2)), CLng(Mid(cap, 9, 2)))
        wLabel = cap & "  (" & Format(d, "d mmm") & " - " & Format(d + 4, "d mmm yyyy") & ")"
    Else
        wLabel = cap
    End If
End Function

Private Sub wShowBox(ByVal ws As Worksheet, sc As SlicerCache)
    Dim box As Object, sh As Object, bp As Object, idx As Long, txt As String, fs As Single
    On Error Resume Next
    Set sh = sc.Slicers(1).Shape
    Set box = ws.Shapes(W_BOX_NAME)
    On Error GoTo 0
    Set bp = wButton(ws, False)

    ' create the box once (clears the old cell text in J4)
    If box Is Nothing Then
        On Error Resume Next
        ws.Range(W_BOX_CELL).Validation.Delete
        ws.Range(W_BOX_CELL).Clear
        On Error GoTo 0
        Set box = ws.Shapes.AddShape(msoShapeRectangle, 0, 0, 100, 30)
        box.Name = W_BOX_NAME
    End If

    ' same look as the < button: rectangle, fill, outline (copied from the button)
    On Error Resume Next
    box.AutoShapeType = msoShapeRectangle
    If Not bp Is Nothing Then
        bp.PickUp
        box.Apply
    Else
        box.Fill.ForeColor.RGB = W_HEADER
        box.Line.ForeColor.RGB = W_LINE
        box.Line.Weight = 1.5
    End If
    box.Placement = xlFreeFloating
    box.OnAction = ""
    On Error GoTo 0

    ' text
    idx = wSelected(sc)
    If idx = 0 Then
        txt = "All - pick one"
    Else
        txt = wLabel(sc.SlicerCacheLevels(1).SlicerItems(idx).Caption)
    End If

    ' position: always J4 (does not follow the slicer); height = button height
    box.Left = ws.Range(W_BOX_CELL).Left
    box.Top = ws.Range(W_BOX_CELL).Top
    If Not bp Is Nothing Then
        box.Height = bp.Height
    Else
        box.Height = ws.Range(W_BOX_CELL).Height
    End If

    ' width: same as the slicer (or one button, see SETTINGS)
    If W_BOX_WIDTH = "button" And Not bp Is Nothing Then
        box.Width = bp.Width
    ElseIf Not sh Is Nothing Then
        box.Width = sh.Width
    End If

    ' text size: as big as SETTINGS allows, but always fits on one line
    fs = W_FONT_SIZE
    If fs > box.Height * 0.6 Then fs = Int(box.Height * 0.6)
    If Len(txt) > 0 Then
        If fs > (box.Width - 8) / (Len(txt) * 0.55) Then fs = Int((box.Width - 8) / (Len(txt) * 0.55))
    End If
    If fs < 8 Then fs = 8
    On Error Resume Next
    With box.TextFrame2
        .WordWrap = msoFalse
        .MarginTop = 0: .MarginBottom = 0
        .MarginLeft = 2: .MarginRight = 2
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.Text = txt
        With .TextRange
            .ParagraphFormat.Alignment = msoAlignCenter
            .Font.Bold = msoTrue
            .Font.Size = fs
            .Font.Fill.ForeColor.RGB = W_TEXT_COLOR
        End With
    End With
    On Error GoTo 0

    ' slicer: top-left corner fixed at J6
    If Not sh Is Nothing Then
        On Error Resume Next
        sh.Placement = xlFreeFloating
        sh.Left = ws.Range(W_SLICER_CELL).Left
        sh.Top = ws.Range(W_SLICER_CELL).Top
        On Error GoTo 0
    End If
End Sub


' ===== COLOURS =====
Private Sub wFormat(ByVal ws As Worksheet)
    Dim pt As PivotTable, c As Range, rowRng As Range
    Dim fr As Long, lr As Long, fc As Long, lc As Long, pctCol As Long, r As Long
    Dim v As Variant

    On Error Resume Next
    Set pt = ws.PivotTables(1)
    On Error GoTo 0
    If pt Is Nothing Then Exit Sub

    fr = pt.TableRange1.Row: fc = pt.TableRange1.Column
    lr = fr + pt.TableRange1.Rows.Count - 1
    lc = fc + pt.TableRange1.Columns.Count - 1

    wKeepButtonsSafe ws

    With ws.Range(ws.Cells(fr, fc), ws.Cells(fr + 300, lc + 10))
        .Interior.ColorIndex = xlNone
        .Font.Bold = False
    End With

    For Each c In ws.Range(ws.Cells(fr, fc), ws.Cells(fr, lc))
        c.Interior.Color = W_HEADER
        c.Font.Bold = True
        If InStr(1, CStr(c.Value), "% Cancelled", vbTextCompare) > 0 Then pctCol = c.Column
    Next c

    For r = fr + 1 To lr
        Set rowRng = ws.Range(ws.Cells(r, fc), ws.Cells(r, lc))
        If InStr(1, CStr(ws.Cells(r, fc).Value), "Grand Total", vbTextCompare) > 0 Then
            rowRng.Interior.Color = W_TOTAL
            rowRng.Font.Bold = True
        ElseIf pctCol > 0 Then
            v = ws.Cells(r, pctCol).Value
            If IsNumeric(v) And Not IsEmpty(v) Then
                If v > W_THRESHOLD Then rowRng.Interior.Color = W_ORANGE
            End If
        End If
    Next r
    ws.Range(ws.Cells(fr, fc), ws.Cells(lr, lc)).Columns.AutoFit
    wAlignToNext ws, lc, fr, lr
End Sub

' last column (Cancellation Reason): always wide enough for its text (never squeezed);
' if the > button reaches further right, the column is stretched to end under the > button
Private Sub wAlignToNext(ByVal ws As Worksheet, ByVal lc As Long, ByVal fr As Long, ByVal lr As Long)
    Dim b As Object, Target As Double, k As Long
    With ws.Range(ws.Cells(fr, lc), ws.Cells(lr, lc))
        .WrapText = False                      ' one line per row
        .Columns.AutoFit                       ' width = longest reason
    End With
    ws.Range(ws.Cells(fr, lc), ws.Cells(lr, lc)).EntireRow.AutoFit   ' undo tall rows
    Set b = wButton(ws, True)
    If b Is Nothing Then Exit Sub
    Target = b.Left + b.Width - ws.Cells(fr, lc).Left
    If Target <= ws.Columns(lc).Width Then Exit Sub   ' text is already wider - keep it
    For k = 1 To 2
        If ws.Columns(lc).Width > 0 Then
            ws.Columns(lc).ColumnWidth = Application.Min(255, ws.Columns(lc).ColumnWidth * Target / ws.Columns(lc).Width)
        End If
    Next k
End Sub


' ===== BUTTON HELPERS =====
' finds the < (isNext = False) or > (isNext = True) button by name or by its assigned macro
Private Function wButton(ByVal ws As Worksheet, ByVal isNext As Boolean) As Object
    Dim shp As Object, act As String, nm As String, keys As Variant, k As Variant
    If isNext Then
        nm = "btnNext": keys = Array("__Next", "NextPeriod", "MonthlyNext")
    Else
        nm = "btnPrev": keys = Array("__Previous", "PreviousPeriod", "MonthlyPrevious")
    End If
    On Error Resume Next
    Set wButton = ws.Shapes(nm)
    On Error GoTo 0
    If Not wButton Is Nothing Then Exit Function
    For Each shp In ws.Shapes
        act = ""
        On Error Resume Next
        act = shp.OnAction
        On Error GoTo 0
        If act <> "" Then
            For Each k In keys
                If InStr(1, act, CStr(k), vbTextCompare) > 0 Then Set wButton = shp: Exit Function
            Next k
        End If
    Next shp
End Function

' buttons do not move/resize with cells; a squashed > button gets its size back
Private Sub wKeepButtonsSafe(ByVal ws As Worksheet)
    Dim bp As Object, bn As Object
    Set bp = wButton(ws, False)
    Set bn = wButton(ws, True)
    On Error Resume Next
    If Not bp Is Nothing Then bp.Placement = xlFreeFloating
    If Not bn Is Nothing Then
        bn.Placement = xlFreeFloating
        If Not bp Is Nothing Then
            If bn.Width < 20 Or bn.Height < 10 Then
                bn.Width = bp.Width
                bn.Height = bp.Height
                bn.Top = bp.Top
                bn.Left = bp.Left + bp.Width
            End If
        End If
    End If
    On Error GoTo 0
End Sub

