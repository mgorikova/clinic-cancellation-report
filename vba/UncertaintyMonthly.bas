Attribute VB_Name = "UncertaintyMonthly"
Option Explicit

' =====================================================================
'  UNCERTAINTY MONTHLY - controls ONLY the "Uncertaintly_Monthly" sheet
'  (module name: UncertaintyMonthly)
'
'  RUN THIS:   UncertMonthly__Main       -> copies the Monthly look, jumps to the latest month,
'                                          colours the table by Signal, draws the month box
'  BUTTONS:    <  = UncertMonthly__Previous      >  = UncertMonthly__Next
'  IF NEEDED:  UncertMonthly__Recolour   (colours look wrong)
'              UncertMonthly__FixButtons (a button got squashed)
'              UncertMonthly__Check      (nothing happens - shows what the macro sees)
'
'  ROW COLOURS (from the Signal column):
'     OK    = light green
'     Low   = light green  (lower than the clinic - change UM_LOW to tell it apart)
'     Alert = light yellow (outside the 95% limit)
'     Alarm = light red    (outside the 99.8% limit)
'
'  Change anything in SETTINGS below - it affects the Uncertaintly_Monthly sheet only
' =====================================================================

' ===== SETTINGS (change these) =====
Private Const UM_SHEET As String = "Uncertaintly_Monthly" ' sheet name
Private Const UM_SLICER As String = "Slicer_Month_Name"  ' slicer (if not found, uses the slicer on the sheet)
Private Const UM_YEAR As String = "2026"                 ' only months of this year ("" = all years)
Private Const UM_BOX_CELL As String = "K4"               ' month box always sits here
Private Const UM_BOX_NAME As String = "lblPeriod"        ' name of the month box shape
Private Const UM_BOX_WIDTH As String = "slicer"          ' month box width: "slicer" or "button"
Private Const UM_SLICER_CELL As String = "K6"            ' slicer top-left corner stays here
Private Const UM_STYLE_FROM As String = "Monthly 2026"    ' copy button look/size + top row heights from this sheet ("" = off)
Private Const UM_FONT_SIZE As Long = 20                  ' max text size in the month box
Private Const UM_TEXT_COLOR As Long = 0                  ' month box text colour (0 = black, 16777215 = white)
Private Const UM_SIGNAL As String = "Signal"             ' header of the Signal column
Private Const UM_HEADER As Long = 16115137               ' header colour        = RGB(193,229,245) #C1E5F5
Private Const UM_OK As Long = 13561798                   ' OK row               = RGB(198,239,206) light green
Private Const UM_LOW As Long = 13561798                  ' Low row              = same green (15523805 = light blue)
Private Const UM_ALERT As Long = 13434879                ' Alert row            = RGB(255,255,204) light yellow
Private Const UM_ALARM As Long = 13551615                ' Alarm row            = RGB(255,199,206) light red
Private Const UM_TOTAL As Long = 13431551                ' Grand Total row      = RGB(255,242,204) #FFF2CC
Private Const UM_LINE As Long = 6567967                  ' box outline (navy)   = RGB(31,56,100)


' ===== MACROS TO RUN =====
Sub UncertMonthly__Main()
    Application.EnableEvents = True
    umCopyStyle
    umMove 0
End Sub

Sub UncertMonthly__Previous()
    umMove -1
End Sub

Sub UncertMonthly__Next()
    umMove 1
End Sub

Sub UncertMonthly__Recolour()
    ' colours + month box ALWAYS update, even if something goes wrong with the buttons;
    ' if a step fails you get a message saying which one (instead of nothing happening)
    Dim ws As Worksheet, sc As SlicerCache, stepName As String
    On Error GoTo Failed
    Set ws = umSheet()
    If ws Is Nothing Then Exit Sub
    Application.ScreenUpdating = False

    stepName = "buttons"
    On Error Resume Next
    umCopyStyle                         ' a button problem never stops the rest
    On Error GoTo Failed

    stepName = "colours"
    umFormat ws

    stepName = "month box"
    Set sc = umSlicer(ws)
    If Not sc Is Nothing Then umShowBox ws, sc

    Application.ScreenUpdating = True
    Exit Sub

Failed:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "UncertMonthly__Recolour stopped at step: " & stepName & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation
End Sub

Sub UncertMonthly__FixButtons()
    Dim ws As Worksheet
    Set ws = umSheet()
    If ws Is Nothing Then Exit Sub
    umKeepButtonsSafe ws
    UncertMonthly__Recolour
End Sub

Sub UncertMonthly__Check()
    Dim ws As Worksheet, sc As SlicerCache, s As String, i As Long
    Dim shp As Object, act As String
    Set ws = umSheet()
    If ws Is Nothing Then MsgBox "Sheet '" & UM_SHEET & "' not found. Check the tab name.": Exit Sub
    s = "Sheet: " & ws.Name & vbCrLf & "Pivots on sheet: " & ws.PivotTables.Count & vbCrLf
    Set sc = umSlicer(ws)
    If sc Is Nothing Then MsgBox s & "NO slicer on this sheet.": Exit Sub
    s = s & "Slicer used: " & sc.Name & "  (" & sc.SourceName & ")" & vbCrLf
    s = s & "First item: " & umFirstCaption(sc) & vbCrLf
    s = s & "Items: " & sc.SlicerCacheLevels(1).SlicerItems.Count & vbCrLf
    s = s & "Selected index: " & umSelected(sc) & vbCrLf
    i = umLatest(sc)
    s = s & "Latest index: " & i
    If i > 0 Then s = s & " = " & sc.SlicerCacheLevels(1).SlicerItems(i).Caption
    s = s & vbCrLf & "< button found: " & CStr(Not umButton(ws, False) Is Nothing)
    s = s & vbCrLf & "> button found: " & CStr(Not umButton(ws, True) Is Nothing)
    s = s & vbCrLf & "Automatic updates (events) on: " & CStr(Application.EnableEvents)
    s = s & vbCrLf & vbCrLf & "SHAPES ON THIS SHEET (name | macro | left, top | width x height | visible):"
    For Each shp In ws.Shapes
        act = ""
        On Error Resume Next
        act = shp.OnAction
        On Error GoTo 0
        s = s & vbCrLf & shp.Name & " | " & act & " | " & Round(shp.Left) & ", " & Round(shp.Top) & _
            " | " & Round(shp.Width) & " x " & Round(shp.Height) & " | " & CStr(shp.Visible = msoTrue)
    Next shp
    MsgBox s
End Sub


' ===== BUTTONS: SAME STYLE AS MONTHLY 2026 (helpers - not run by hand) =====
' - if a < or > button is missing, it is made again (assigned to the right macro)
' - buttons are never hidden, never squashed, always in front
' - look + size copied from the Monthly 2026 buttons
' - position: left where it is, EXCEPT when the button is gone from the top area
'   (off the sheet, under the table, or squashed) -> put back where Monthly 2026 has it
Private Sub umCopyStyle()
    Dim ws As Worksheet, src As Worksheet
    Dim sp As Object, sn As Object, bp As Object, bn As Object
    Dim r As Long, lastRow As Long, topLimit As Double

    Set ws = umSheet()
    If ws Is Nothing Then Exit Sub
    If UM_STYLE_FROM <> "" Then
        On Error Resume Next
        Set src = ThisWorkbook.Worksheets(UM_STYLE_FROM)
        On Error GoTo 0
    End If
    If Not src Is Nothing Then
        Set sp = umButton(src, False): Set sn = umButton(src, True)
    End If

    ' 0) find the buttons - make them again if they are missing
    Set bp = umButton(ws, False)
    Set bn = umButton(ws, True)
    If bp Is Nothing Then Set bp = umMakeButton(ws, False, sp)
    If bn Is Nothing Then Set bn = umMakeButton(ws, True, sn)

    ' 1) never move/shrink with rows or columns, never hidden (set BEFORE row heights change)
    On Error Resume Next
    bp.Placement = xlFreeFloating: bn.Placement = xlFreeFloating
    bp.Visible = msoTrue:          bn.Visible = msoTrue
    bp.OnAction = "UncertMonthly__Previous"
    bn.OnAction = "UncertMonthly__Next"
    On Error GoTo 0

    ' 2) same row heights above the slicer as Monthly 2026
    If Not src Is Nothing Then
        lastRow = ws.Range(UM_SLICER_CELL).Row - 1
        For r = 1 To lastRow
            ws.Rows(r).Hidden = False
            ws.Rows(r).RowHeight = src.Rows(r).RowHeight
        Next r
    End If

    ' 3) same look + size as the Monthly 2026 buttons
    On Error Resume Next
    If Not sp Is Nothing Then
        bp.AutoShapeType = sp.AutoShapeType
        sp.PickUp: bp.Apply
        bp.TextFrame2.TextRange.Font.Size = sp.TextFrame2.TextRange.Font.Size
        bp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = sp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
        bp.LockAspectRatio = msoFalse
        bp.Width = sp.Width: bp.Height = sp.Height
    End If
    If Not sn Is Nothing Then
        bn.AutoShapeType = sn.AutoShapeType
        sn.PickUp: bn.Apply
        bn.TextFrame2.TextRange.Font.Size = sn.TextFrame2.TextRange.Font.Size
        bn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = sn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
        bn.LockAspectRatio = msoFalse
        bn.Width = sn.Width: bn.Height = sn.Height
    End If
    If bp.Width < 20 Or bp.Height < 10 Then bp.Width = 90: bp.Height = 30
    If bn.Width < 20 Or bn.Height < 10 Then bn.Width = bp.Width: bn.Height = bp.Height
    On Error GoTo 0

    ' 4) a button that left the top area goes back to the Monthly 2026 position
    topLimit = ws.Range(UM_SLICER_CELL).Top
    If bp.Top < 0 Or bp.Left < 0 Or bp.Top + bp.Height > topLimit + 2 Then
        If Not sp Is Nothing Then
            bp.Left = sp.Left: bp.Top = sp.Top
        Else
            bp.Left = ws.Range(UM_BOX_CELL).Left - 2 * (bp.Width + 6): bp.Top = ws.Range(UM_BOX_CELL).Top
        End If
    End If
    If bn.Top < 0 Or bn.Left < 0 Or bn.Top + bn.Height > topLimit + 2 Or _
       Abs(bn.Top - bp.Top) > 1 Or bn.Left < bp.Left + bp.Width Then
        If Not sn Is Nothing And Not sp Is Nothing Then
            bn.Top = bp.Top
            bn.Left = bp.Left + bp.Width + (sn.Left - (sp.Left + sp.Width))
        Else
            bn.Top = bp.Top
            bn.Left = bp.Left + bp.Width + 6
        End If
    End If

    bp.ZOrder msoBringToFront: bn.ZOrder msoBringToFront
End Sub

' makes a new < or > button (only used when one is missing)
Private Function umMakeButton(ByVal ws As Worksheet, ByVal isNext As Boolean, ByVal model As Object) As Object
    Dim b As Object
    Set b = ws.Shapes.AddShape(msoShapeRectangle, 0, 0, 90, 30)
    If isNext Then
        b.Name = "btnNext"
        b.TextFrame2.TextRange.Text = ">"
        b.OnAction = "UncertMonthly__Next"
    Else
        b.Name = "btnPrev"
        b.TextFrame2.TextRange.Text = "<"
        b.OnAction = "UncertMonthly__Previous"
    End If
    b.Fill.ForeColor.RGB = UM_HEADER
    b.Line.ForeColor.RGB = UM_LINE
    b.Line.Weight = 1.5
    b.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    b.TextFrame2.VerticalAnchor = msoAnchorMiddle
    If Not model Is Nothing Then
        b.Left = model.Left: b.Top = model.Top
        b.Width = model.Width: b.Height = model.Height
    Else
        b.Top = ws.Range(UM_BOX_CELL).Top
        b.Left = ws.Range(UM_BOX_CELL).Left - IIf(isNext, 1, 2) * (b.Width + 6)
    End If
    Set umMakeButton = b
End Function


' ===== NAVIGATION (helpers - not run by hand) =====
Private Function umSheet() As Worksheet
    On Error Resume Next
    Set umSheet = ThisWorkbook.Worksheets(UM_SHEET)
    On Error GoTo 0
End Function

Private Function umSlicer(ByVal ws As Worksheet) As SlicerCache
    ' 1st: slicer on this sheet with items like "01-January" or "2026-01 | January"
    ' 2nd: slicer named in SETTINGS (if on this sheet)   3rd: any slicer on this sheet
    Dim sc As SlicerCache, firstOnSheet As SlicerCache, named As SlicerCache
    For Each sc In ThisWorkbook.SlicerCaches
        If umOnSheet(sc, ws) Then
            If firstOnSheet Is Nothing Then Set firstOnSheet = sc
            If sc.Name = UM_SLICER Then Set named = sc
            If umFirstCaption(sc) Like "##-*" Or umFirstCaption(sc) Like "####-## | *" Then Set umSlicer = sc: Exit Function
        End If
    Next sc
    If Not named Is Nothing Then Set umSlicer = named: Exit Function
    Set umSlicer = firstOnSheet
End Function

Private Function umOnSheet(sc As SlicerCache, ByVal ws As Worksheet) As Boolean
    Dim i As Long
    On Error Resume Next
    For i = 1 To sc.Slicers.Count
        If sc.Slicers(i).Shape.Parent.Name = ws.Name Then umOnSheet = True: Exit Function
    Next i
End Function

Private Function umFirstCaption(sc As SlicerCache) As String
    Dim itms As SlicerItems, i As Long
    On Error Resume Next
    Set itms = sc.SlicerCacheLevels(1).SlicerItems
    If itms Is Nothing Then Exit Function
    For i = 1 To itms.Count
        If Trim(itms(i).Caption) <> "" Then umFirstCaption = itms(i).Caption: Exit Function
    Next i
End Function

Private Function umInScope(it As SlicerItem) As Boolean
    If Not it.HasData Then Exit Function
    If Not (it.Caption Like "#*") Then Exit Function      ' skip "(blank)" and other non-date items
    If UM_YEAR <> "" And it.Caption Like "####*" Then      ' year filter only if caption starts with a year
        If Left(it.Caption, 4) <> UM_YEAR Then Exit Function
    End If
    umInScope = True
End Function

Private Function umSelected(sc As SlicerCache) As Long
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
        If itms(i).Name = cur Then umSelected = i: Exit Function
    Next i
End Function

Private Function umStep(sc As SlicerCache, ByVal idx As Long, ByVal stepBy As Long) As Long
    ' Previous = closest EARLIER month, Next = closest LATER month.
    ' Compares the month names themselves ("01-January" < "02-February"), so it works whatever order the slicer shows them in.
    Dim itms As SlicerItems, i As Long, cur As String, cap As String, best As String
    Set itms = sc.SlicerCacheLevels(1).SlicerItems
    cur = itms(idx).Caption
    For i = 1 To itms.Count
        If umInScope(itms(i)) Then
            cap = itms(i).Caption
            If stepBy < 0 Then
                If cap < cur Then
                    If umStep = 0 Or cap > best Then umStep = i: best = cap
                End If
            Else
                If cap > cur Then
                    If umStep = 0 Or cap < best Then umStep = i: best = cap
                End If
            End If
        End If
    Next i
End Function

Private Function umLatest(sc As SlicerCache) As Long
    ' the newest month that has data (and is in the year from SETTINGS)
    Dim itms As SlicerItems, i As Long, best As String
    Set itms = sc.SlicerCacheLevels(1).SlicerItems
    For i = 1 To itms.Count
        If umInScope(itms(i)) Then
            If umLatest = 0 Or itms(i).Caption > best Then umLatest = i: best = itms(i).Caption
        End If
    Next i
End Function

Private Sub umMove(ByVal stepBy As Long)
    Dim ws As Worksheet, sc As SlicerCache, idx As Long, i As Long
    Set ws = umSheet()
    If ws Is Nothing Then MsgBox "Sheet '" & UM_SHEET & "' not found.": Exit Sub
    Set sc = umSlicer(ws)
    If sc Is Nothing Then MsgBox "No slicer found on sheet '" & UM_SHEET & "'.": Exit Sub

    idx = umSelected(sc)
    If idx > 0 Then
        If Not umInScope(sc.SlicerCacheLevels(1).SlicerItems(idx)) Then idx = 0
    End If
    If stepBy = 0 Or idx = 0 Then
        i = umLatest(sc)
    Else
        i = umStep(sc, idx, stepBy)
    End If
    If i = 0 Then
        If stepBy = 0 Or idx = 0 Then
            MsgBox "No month found to show." & vbCrLf & _
                   "Slicer used: " & sc.Name & vbCrLf & _
                   "First item: " & umFirstCaption(sc) & vbCrLf & vbCrLf & _
                   "Run UncertMonthly__Check for details."
        Else
            Beep                                            ' already at first / last month
        End If
        Exit Sub
    End If
    umSelect ws, sc, i
End Sub

Private Sub umSelect(ByVal ws As Worksheet, sc As SlicerCache, ByVal i As Long)
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
    umFormat ws
    umShowBox ws, sc
End Sub


' ===== MONTH BOX (fixed at K4, looks like the < > buttons, width = slicer; slicer fixed at K6) =====
Private Function umLabel(ByVal cap As String) As String
    ' "01-January" shown as it is (same as the Monthly 2026 box)
    ' "2026-10 | October" -> "October 2026"
    If cap Like "####-## | *" Then
        umLabel = Trim(Mid(cap, 11)) & " " & Left(cap, 4)
    Else
        umLabel = cap
    End If
End Function

Private Sub umShowBox(ByVal ws As Worksheet, sc As SlicerCache)
    Dim box As Object, sh As Object, bp As Object, idx As Long, txt As String, fs As Single
    On Error Resume Next
    Set sh = sc.Slicers(1).Shape
    Set box = ws.Shapes(UM_BOX_NAME)
    On Error GoTo 0
    Set bp = umButton(ws, False)

    ' create the box once (clears the old cell text in K4)
    If box Is Nothing Then
        On Error Resume Next
        ws.Range(UM_BOX_CELL).Validation.Delete
        ws.Range(UM_BOX_CELL).Clear
        On Error GoTo 0
        Set box = ws.Shapes.AddShape(msoShapeRectangle, 0, 0, 100, 30)
        box.Name = UM_BOX_NAME
    End If

    ' same look as the < button: rectangle, fill, outline (copied from the button)
    On Error Resume Next
    box.AutoShapeType = msoShapeRectangle
    If Not bp Is Nothing Then
        bp.PickUp
        box.Apply
    Else
        box.Fill.ForeColor.RGB = UM_HEADER
        box.Line.ForeColor.RGB = UM_LINE
        box.Line.Weight = 1.5
    End If
    box.Placement = xlFreeFloating
    box.OnAction = ""
    On Error GoTo 0

    ' text
    idx = umSelected(sc)
    If idx = 0 Then
        txt = "All - pick one"
    Else
        txt = umLabel(sc.SlicerCacheLevels(1).SlicerItems(idx).Caption)
    End If

    ' position: always K4 (does not follow the slicer); height = button height
    box.Left = ws.Range(UM_BOX_CELL).Left
    box.Top = ws.Range(UM_BOX_CELL).Top
    If Not bp Is Nothing Then
        box.Height = bp.Height
    Else
        box.Height = ws.Range(UM_BOX_CELL).Height
    End If

    ' width: same as the slicer (or one button, see SETTINGS)
    If UM_BOX_WIDTH = "button" And Not bp Is Nothing Then
        box.Width = bp.Width
    ElseIf Not sh Is Nothing Then
        box.Width = sh.Width
    End If

    ' text size: as big as SETTINGS allows, but always fits on one line
    fs = UM_FONT_SIZE
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
            .Font.Fill.ForeColor.RGB = UM_TEXT_COLOR
        End With
    End With
    On Error GoTo 0

    ' slicer: top-left corner fixed at K6
    If Not sh Is Nothing Then
        On Error Resume Next
        sh.Placement = xlFreeFloating
        sh.Left = ws.Range(UM_SLICER_CELL).Left
        sh.Top = ws.Range(UM_SLICER_CELL).Top
        On Error GoTo 0
    End If
End Sub


' ===== COLOURS =====
Private Sub umFormat(ByVal ws As Worksheet)
    Dim pt As PivotTable, c As Range, rowRng As Range
    Dim fr As Long, lr As Long, fc As Long, lc As Long, sigCol As Long, r As Long

    On Error Resume Next
    Set pt = ws.PivotTables(1)
    On Error GoTo 0
    If pt Is Nothing Then Exit Sub

    fr = pt.TableRange1.Row: fc = pt.TableRange1.Column
    lr = fr + pt.TableRange1.Rows.Count - 1
    lc = fc + pt.TableRange1.Columns.Count - 1

    umKeepButtonsSafe ws

    ' wipe ALL old colours first (also rows left over from a longer month)
    With ws.Range(ws.Cells(fr, fc), ws.Cells(fr + 300, lc + 10))
        .Interior.ColorIndex = xlNone
        .Font.Bold = False
    End With

    For Each c In ws.Range(ws.Cells(fr, fc), ws.Cells(fr, lc))
        c.Interior.Color = UM_HEADER
        c.Font.Bold = True
        If sigCol = 0 And StrComp(Trim(CStr(c.Value)), UM_SIGNAL, vbTextCompare) = 0 Then sigCol = c.Column
    Next c

    For r = fr + 1 To lr
        Set rowRng = ws.Range(ws.Cells(r, fc), ws.Cells(r, lc))
        If InStr(1, CStr(ws.Cells(r, fc).Value), "Grand Total", vbTextCompare) > 0 Then
            rowRng.Interior.Color = UM_TOTAL
            rowRng.Font.Bold = True
        ElseIf sigCol > 0 Then
            Select Case LCase(Trim(CStr(ws.Cells(r, sigCol).Value)))
                Case "ok":    rowRng.Interior.Color = UM_OK
                Case "low":   rowRng.Interior.Color = UM_LOW
                Case "alert": rowRng.Interior.Color = UM_ALERT
                Case "alarm": rowRng.Interior.Color = UM_ALARM
            End Select
        End If
    Next r
    ws.Range(ws.Cells(fr, fc), ws.Cells(lr, lc)).Columns.AutoFit
    umAlignToNext ws, lc, fr, lr
End Sub

' last column (Signal): always wide enough for its text (never squeezed);
' if the > button reaches further right, the column is stretched to end under the > button
Private Sub umAlignToNext(ByVal ws As Worksheet, ByVal lc As Long, ByVal fr As Long, ByVal lr As Long)
    Dim b As Object, Target As Double, k As Long
    With ws.Range(ws.Cells(fr, lc), ws.Cells(lr, lc))
        .WrapText = False                      ' one line per row
        .Columns.AutoFit                       ' width = longest text
    End With
    ws.Range(ws.Cells(fr, lc), ws.Cells(lr, lc)).EntireRow.AutoFit   ' undo tall rows
    Set b = umButton(ws, True)
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
Private Function umButton(ByVal ws As Worksheet, ByVal isNext As Boolean) As Object
    Dim shp As Object, act As String, nm As String, keys As Variant, k As Variant
    If isNext Then
        nm = "btnNext": keys = Array("__Next", "NextPeriod", "MonthlyNext")
    Else
        nm = "btnPrev": keys = Array("__Previous", "PreviousPeriod", "MonthlyPrevious")
    End If
    On Error Resume Next
    Set umButton = ws.Shapes(nm)
    On Error GoTo 0
    If Not umButton Is Nothing Then Exit Function
    For Each shp In ws.Shapes
        act = ""
        On Error Resume Next
        act = shp.OnAction
        On Error GoTo 0
        If act <> "" Then
            For Each k In keys
                If InStr(1, act, CStr(k), vbTextCompare) > 0 Then Set umButton = shp: Exit Function
            Next k
        End If
    Next shp
End Function

' buttons do not move/resize with cells; a squashed > button gets its size back
Private Sub umKeepButtonsSafe(ByVal ws As Worksheet)
    Dim bp As Object, bn As Object
    Set bp = umButton(ws, False)
    Set bn = umButton(ws, True)
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


' =====================================================================
'  AUTOMATIC UPDATE when you click the slicer - SHEET CODE
'  Alt+F11 > double-click the SHEET "Uncertaintly_Monthly" > paste:
'
'  Option Explicit
'  Private Sub Worksheet_PivotTableUpdate(ByVal Target As PivotTable)
'      Application.EnableEvents = False
'      Application.Run "'" & ThisWorkbook.Name & "'!UncertMonthly__Recolour"
'      Application.EnableEvents = True
'  End Sub
' =====================================================================

