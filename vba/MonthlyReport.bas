Attribute VB_Name = "MonthlyReport"
Option Explicit

' =====================================================================
'  MONTHLY REPORT - controls ONLY the "Monthly 2026" sheet
'
'  RUN THIS:   Monthly__Main        -> jumps to the latest month, colours the table, draws the month box
'  BUTTONS:    <  = Monthly__Previous      >  = Monthly__Next
'  IF NEEDED:  Monthly__Recolour   (colours look wrong)
'              Monthly__FixButtons (a button got squashed)
'              Monthly__Check      (nothing happens - shows what the macro sees)
'
'  Change anything in SETTINGS below - it affects the Monthly 2026 sheet only
' =====================================================================

' ===== SETTINGS (change these) =====
Private Const M_SHEET As String = "Monthly 2026"        ' sheet name
Private Const M_SLICER As String = "Slicer_Month_Label" ' slicer (if not found, uses the slicer on the sheet)
Private Const M_YEAR As String = "2026"                 ' only months of this year ("" = all years)
Private Const M_BOX_CELL As String = "J4"               ' month box always sits here
Private Const M_BOX_NAME As String = "lblPeriod"        ' name of the month box shape
Private Const M_BOX_WIDTH As String = "slicer"          ' month box width: "slicer" or "button"
Private Const M_SLICER_CELL As String = "J6"            ' slicer top-left corner stays here
Private Const M_FONT_SIZE As Long = 20                  ' max text size in the month box
Private Const M_TEXT_COLOR As Long = 0                  ' month box text colour (0 = black, 16777215 = white)
Private Const M_THRESHOLD As Double = 0.2               ' orange row if % Cancelled > 20%
Private Const M_HEADER As Long = 16115137               ' header colour        = RGB(193,229,245) #C1E5F5
Private Const M_ORANGE As Long = 11851260               ' high cancel row      = RGB(252,213,180)
Private Const M_TOTAL As Long = 13431551                ' Grand Total row      = RGB(255,242,204) #FFF2CC
Private Const M_LINE As Long = 6567967                  ' box outline (navy)   = RGB(31,56,100)


' ===== MACROS TO RUN =====
Sub Monthly__Main()
    mMove 0
End Sub

Sub Monthly__Previous()
    mMove -1
End Sub

Sub Monthly__Next()
    mMove 1
End Sub

Sub Monthly__Recolour()
    Dim ws As Worksheet, sc As SlicerCache
    Set ws = mSheet()
    If ws Is Nothing Then Exit Sub
    mFormat ws
    Set sc = mSlicer(ws)
    If Not sc Is Nothing Then mShowBox ws, sc
End Sub

Sub Monthly__FixButtons()
    Dim ws As Worksheet
    Set ws = mSheet()
    If ws Is Nothing Then Exit Sub
    mKeepButtonsSafe ws
    Monthly__Recolour
End Sub

Sub Monthly__Check()
    Dim ws As Worksheet, sc As SlicerCache, s As String, i As Long
    Set ws = mSheet()
    If ws Is Nothing Then MsgBox "Sheet '" & M_SHEET & "' not found. Check the tab name.": Exit Sub
    s = "Sheet: " & ws.Name & vbCrLf & "Pivots on sheet: " & ws.PivotTables.Count & vbCrLf
    Set sc = mSlicer(ws)
    If sc Is Nothing Then MsgBox s & "NO slicer on this sheet.": Exit Sub
    s = s & "Slicer used: " & sc.Name & "  (" & sc.SourceName & ")" & vbCrLf
    s = s & "First item: " & mFirstCaption(sc) & vbCrLf
    s = s & "Items: " & sc.SlicerCacheLevels(1).SlicerItems.Count & vbCrLf
    s = s & "Selected index: " & mSelected(sc) & vbCrLf
    i = mLatest(sc)
    s = s & "Latest index: " & i
    If i > 0 Then s = s & " = " & sc.SlicerCacheLevels(1).SlicerItems(i).Caption
    s = s & vbCrLf & "< button found: " & CStr(Not mButton(ws, False) Is Nothing)
    s = s & vbCrLf & "> button found: " & CStr(Not mButton(ws, True) Is Nothing)
    MsgBox s
End Sub


' ===== NAVIGATION (helpers - not run by hand) =====
Private Function mSheet() As Worksheet
    On Error Resume Next
    Set mSheet = ThisWorkbook.Worksheets(M_SHEET)
    On Error GoTo 0
End Function

Private Function mSlicer(ByVal ws As Worksheet) As SlicerCache
    ' 1st: slicer on this sheet with items like "2026-01 | January"
    ' 2nd: slicer named in SETTINGS (if on this sheet)   3rd: any slicer on this sheet
    Dim sc As SlicerCache, firstOnSheet As SlicerCache, named As SlicerCache
    For Each sc In ThisWorkbook.SlicerCaches
        If mOnSheet(sc, ws) Then
            If firstOnSheet Is Nothing Then Set firstOnSheet = sc
            If sc.Name = M_SLICER Then Set named = sc
            If mFirstCaption(sc) Like "####-## | *" Then Set mSlicer = sc: Exit Function
        End If
    Next sc
    If Not named Is Nothing Then Set mSlicer = named: Exit Function
    Set mSlicer = firstOnSheet
End Function

Private Function mOnSheet(sc As SlicerCache, ByVal ws As Worksheet) As Boolean
    Dim i As Long
    On Error Resume Next
    For i = 1 To sc.Slicers.Count
        If sc.Slicers(i).Shape.Parent.Name = ws.Name Then mOnSheet = True: Exit Function
    Next i
End Function

Private Function mFirstCaption(sc As SlicerCache) As String
    Dim itms As SlicerItems, i As Long
    On Error Resume Next
    Set itms = sc.SlicerCacheLevels(1).SlicerItems
    If itms Is Nothing Then Exit Function
    For i = 1 To itms.Count
        If Trim(itms(i).Caption) <> "" Then mFirstCaption = itms(i).Caption: Exit Function
    Next i
End Function

Private Function mInScope(it As SlicerItem) As Boolean
    If Not it.HasData Then Exit Function
    If Not (it.Caption Like "#*") Then Exit Function      ' skip "(blank)" and other non-date items
    If M_YEAR <> "" And it.Caption Like "####*" Then      ' year filter only if caption starts with a year
        If Left(it.Caption, 4) <> M_YEAR Then Exit Function
    End If
    mInScope = True
End Function

Private Function mSelected(sc As SlicerCache) As Long
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
        If itms(i).Name = cur Then mSelected = i: Exit Function
    Next i
End Function

Private Function mStep(sc As SlicerCache, ByVal idx As Long, ByVal stepBy As Long) As Long
    ' Previous = closest EARLIER month, Next = closest LATER month.
    ' Compares the month names themselves, so it works whatever order the slicer shows them in.
    Dim itms As SlicerItems, i As Long, cur As String, cap As String, best As String
    Set itms = sc.SlicerCacheLevels(1).SlicerItems
    cur = itms(idx).Caption
    For i = 1 To itms.Count
        If mInScope(itms(i)) Then
            cap = itms(i).Caption
            If stepBy < 0 Then
                If cap < cur Then
                    If mStep = 0 Or cap > best Then mStep = i: best = cap
                End If
            Else
                If cap > cur Then
                    If mStep = 0 Or cap < best Then mStep = i: best = cap
                End If
            End If
        End If
    Next i
End Function

Private Function mLatest(sc As SlicerCache) As Long
    ' the newest month that has data (and is in the year from SETTINGS)
    Dim itms As SlicerItems, i As Long, best As String
    Set itms = sc.SlicerCacheLevels(1).SlicerItems
    For i = 1 To itms.Count
        If mInScope(itms(i)) Then
            If mLatest = 0 Or itms(i).Caption > best Then mLatest = i: best = itms(i).Caption
        End If
    Next i
End Function

Private Sub mMove(ByVal stepBy As Long)
    Dim ws As Worksheet, sc As SlicerCache, idx As Long, i As Long
    Set ws = mSheet()
    If ws Is Nothing Then MsgBox "Sheet '" & M_SHEET & "' not found.": Exit Sub
    Set sc = mSlicer(ws)
    If sc Is Nothing Then MsgBox "No slicer found on sheet '" & M_SHEET & "'.": Exit Sub

    idx = mSelected(sc)
    If idx > 0 Then
        If Not mInScope(sc.SlicerCacheLevels(1).SlicerItems(idx)) Then idx = 0
    End If
    If stepBy = 0 Or idx = 0 Then
        i = mLatest(sc)
    Else
        i = mStep(sc, idx, stepBy)
    End If
    If i = 0 Then
        If stepBy = 0 Or idx = 0 Then
            MsgBox "No month found to show." & vbCrLf & _
                   "Slicer used: " & sc.Name & vbCrLf & _
                   "First item: " & mFirstCaption(sc) & vbCrLf & vbCrLf & _
                   "Run Monthly__Check for details."
        Else
            Beep                                            ' already at first / last month
        End If
        Exit Sub
    End If
    mSelect ws, sc, i
End Sub

Private Sub mSelect(ByVal ws As Worksheet, sc As SlicerCache, ByVal i As Long)
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
    mFormat ws
    mShowBox ws, sc
End Sub


' ===== MONTH BOX (fixed at J4, looks like the < > buttons, width = slicer; slicer fixed at J6) =====
Private Function mLabel(ByVal cap As String) As String
    ' "2026-10 | October" -> "October 2026"   (other captions, e.g. "10-October", shown as they are)
    If cap Like "####-## | *" Then
        mLabel = Trim(Mid(cap, 11)) & " " & Left(cap, 4)
    Else
        mLabel = cap
    End If
End Function

Private Sub mShowBox(ByVal ws As Worksheet, sc As SlicerCache)
    Dim box As Object, sh As Object, bp As Object, idx As Long, txt As String, fs As Single
    On Error Resume Next
    Set sh = sc.Slicers(1).Shape
    Set box = ws.Shapes(M_BOX_NAME)
    On Error GoTo 0
    Set bp = mButton(ws, False)

    ' create the box once (clears the old cell text in J4)
    If box Is Nothing Then
        On Error Resume Next
        ws.Range(M_BOX_CELL).Validation.Delete
        ws.Range(M_BOX_CELL).Clear
        On Error GoTo 0
        Set box = ws.Shapes.AddShape(msoShapeRectangle, 0, 0, 100, 30)
        box.Name = M_BOX_NAME
    End If

    ' same look as the < button: rectangle, fill, outline (copied from the button)
    On Error Resume Next
    box.AutoShapeType = msoShapeRectangle
    If Not bp Is Nothing Then
        bp.PickUp
        box.Apply
    Else
        box.Fill.ForeColor.RGB = M_HEADER
        box.Line.ForeColor.RGB = M_LINE
        box.Line.Weight = 1.5
    End If
    box.Placement = xlFreeFloating
    box.OnAction = ""
    On Error GoTo 0

    ' text
    idx = mSelected(sc)
    If idx = 0 Then
        txt = "All - pick one"
    Else
        txt = mLabel(sc.SlicerCacheLevels(1).SlicerItems(idx).Caption)
    End If

    ' position: always J4 (does not follow the slicer); height = button height
    box.Left = ws.Range(M_BOX_CELL).Left
    box.Top = ws.Range(M_BOX_CELL).Top
    If Not bp Is Nothing Then
        box.Height = bp.Height
    Else
        box.Height = ws.Range(M_BOX_CELL).Height
    End If

    ' width: same as the slicer (or one button, see SETTINGS)
    If M_BOX_WIDTH = "button" And Not bp Is Nothing Then
        box.Width = bp.Width
    ElseIf Not sh Is Nothing Then
        box.Width = sh.Width
    End If

    ' text size: as big as SETTINGS allows, but always fits on one line
    fs = M_FONT_SIZE
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
            .Font.Fill.ForeColor.RGB = M_TEXT_COLOR
        End With
    End With
    On Error GoTo 0

    ' slicer: top-left corner fixed at J6
    If Not sh Is Nothing Then
        On Error Resume Next
        sh.Placement = xlFreeFloating
        sh.Left = ws.Range(M_SLICER_CELL).Left
        sh.Top = ws.Range(M_SLICER_CELL).Top
        On Error GoTo 0
    End If
End Sub


' ===== COLOURS =====
Private Sub mFormat(ByVal ws As Worksheet)
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

    mKeepButtonsSafe ws

    With ws.Range(ws.Cells(fr, fc), ws.Cells(fr + 300, lc + 10))
        .Interior.ColorIndex = xlNone
        .Font.Bold = False
    End With

    For Each c In ws.Range(ws.Cells(fr, fc), ws.Cells(fr, lc))
        c.Interior.Color = M_HEADER
        c.Font.Bold = True
        If InStr(1, CStr(c.Value), "% Cancelled", vbTextCompare) > 0 Then pctCol = c.Column
    Next c

    For r = fr + 1 To lr
        Set rowRng = ws.Range(ws.Cells(r, fc), ws.Cells(r, lc))
        If InStr(1, CStr(ws.Cells(r, fc).Value), "Grand Total", vbTextCompare) > 0 Then
            rowRng.Interior.Color = M_TOTAL
            rowRng.Font.Bold = True
        ElseIf pctCol > 0 Then
            v = ws.Cells(r, pctCol).Value
            If IsNumeric(v) And Not IsEmpty(v) Then
                If v > M_THRESHOLD Then rowRng.Interior.Color = M_ORANGE
            End If
        End If
    Next r
    ws.Range(ws.Cells(fr, fc), ws.Cells(lr, lc)).Columns.AutoFit
    mAlignToNext ws, lc, fr, lr
End Sub

' last column (Cancellation Reason) ends under the right edge of the > button
Private Sub mAlignToNext(ByVal ws As Worksheet, ByVal lc As Long, ByVal fr As Long, ByVal lr As Long)
    Dim b As Object, Target As Double, k As Long
    Set b = mButton(ws, True)
    If b Is Nothing Then Exit Sub
    Target = b.Left + b.Width - ws.Cells(fr, lc).Left
    If Target < 20 Then Exit Sub
    For k = 1 To 2
        If ws.Columns(lc).Width > 0 Then
            ws.Columns(lc).ColumnWidth = Application.Min(255, ws.Columns(lc).ColumnWidth * Target / ws.Columns(lc).Width)
        End If
    Next k
    ws.Range(ws.Cells(fr + 1, lc), ws.Cells(lr, lc)).WrapText = True
End Sub


' ===== BUTTON HELPERS =====
' finds the < (isNext = False) or > (isNext = True) button by name or by its assigned macro
Private Function mButton(ByVal ws As Worksheet, ByVal isNext As Boolean) As Object
    Dim shp As Object, act As String, nm As String, keys As Variant, k As Variant
    If isNext Then
        nm = "btnNext": keys = Array("__Next", "NextPeriod", "MonthlyNext")
    Else
        nm = "btnPrev": keys = Array("__Previous", "PreviousPeriod", "MonthlyPrevious")
    End If
    On Error Resume Next
    Set mButton = ws.Shapes(nm)
    On Error GoTo 0
    If Not mButton Is Nothing Then Exit Function
    For Each shp In ws.Shapes
        act = ""
        On Error Resume Next
        act = shp.OnAction
        On Error GoTo 0
        If act <> "" Then
            For Each k In keys
                If InStr(1, act, CStr(k), vbTextCompare) > 0 Then Set mButton = shp: Exit Function
            Next k
        End If
    Next shp
End Function

' buttons do not move/resize with cells; a squashed > button gets its size back
Private Sub mKeepButtonsSafe(ByVal ws As Worksheet)
    Dim bp As Object, bn As Object
    Set bp = mButton(ws, False)
    Set bn = mButton(ws, True)
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

