Attribute VB_Name = "modMain"
Option Explicit

' Import an OpenAPI YAML file into this workbook.
'
' One sheet is made per operation from the hidden template sheet
' "_tpl_endpoint". When a sheet for the same operationId already exists it
' is replaced, but the rows between @processing:begin and @processing:end
' (the part people write by hand) are moved over to the new sheet.

Public Const TPL_ENDPOINT As String = "_tpl_endpoint"
Public Const TPL_INDEX As String = "_tpl_index"
Public Const CONFIG_SHEET As String = "_config"
Private Const KEY_PREFIX As String = "opid:"
Private Const INDEX_KEY As String = "@index"

' Where the import is, for error messages.
Private mStep As String

' Called from the button on the guide sheet.
Public Sub ImportOpenApi()
    Dim path As String, result As String
    path = PickYamlFile()
    If path = "" Then Exit Sub
    result = ImportFromPath(path)
    If Left$(result, 2) = "OK" Then
        MsgBox Label("msgDone") & vbLf & Mid$(result, 4), vbInformation
    Else
        MsgBox Label("msgFailed") & vbLf & vbLf & result, vbExclamation
    End If
End Sub

' Import without dialogs. Returns "OK <message>" or "ERROR: <message>".
Public Function ImportFromPath(ByVal path As String) As String
    Dim root As clsNode, model As clsNode, count As Long
    Dim oldScreen As Boolean, oldAlerts As Boolean, oldEvents As Boolean, oldCalc As Long

    On Error GoTo Failed
    mStep = "read file"
    Set root = YamlParse(ReadUtf8File(path))
    mStep = "read OpenAPI"
    Set model = BuildModel(root, path)

    mStep = "backup"
    If Label("backup") = "on" Then BackupWorkbook

    oldScreen = Application.ScreenUpdating
    oldAlerts = Application.DisplayAlerts
    oldEvents = Application.EnableEvents
    oldCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    count = WriteSheets(model)
    SetConfig "lastPath", path
    SetConfig "lastImport", model.Text("importedAt")

    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts
    Application.ScreenUpdating = oldScreen
    ImportFromPath = "OK " & count & " API"
    Exit Function

Failed:
    ImportFromPath = "ERROR: " & Err.Description & " (" & mStep & ", error " & Err.Number & ")"
    On Error Resume Next
    If oldCalc <> 0 Then Application.Calculation = oldCalc
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Function

' ---------------------------------------------------------------------------
' Sheets
' ---------------------------------------------------------------------------

Private Function WriteSheets(ByVal model As clsNode) As Long
    Dim ops As clsNode, op As clsNode, indexRows As clsNode, row As clsNode
    Dim ws As Worksheet, old As Worksheet, prev As Worksheet, indexWs As Worksheet, first As Worksheet
    Dim i As Long, names As clsNode, key As String

    Set ops = model.Child("operations")
    Set names = NewMap()   ' operationId -> sheet name
    Set indexRows = NewSeq()

    For i = 1 To ops.Count
        Set op = ops.ItemAt(i)
        mStep = "sheet " & op.Text("operationId") & ": copy template"
        Set old = FindSheetByKey(KEY_PREFIX & op.Text("operationId"))
        Set ws = CopyTemplate(TPL_ENDPOINT)
        mStep = "sheet " & op.Text("operationId") & ": fill"
        FillSheet ws, NewCtx(model, op)
        If Not old Is Nothing Then
            mStep = "sheet " & op.Text("operationId") & ": move processing"
            MoveProcessing old, ws
            ws.Move Before:=old
            old.Delete
        End If
        mStep = "sheet " & op.Text("operationId") & ": rename"
        ws.Name = UniqueSheetName(op.Text("operationId"))
        ws.Tab.ColorIndex = xlColorIndexNone
        names.SetText op.Text("operationId"), ws.Name

        Set row = NewMap()
        row.SetText "no", CStr(i)
        row.SetText "operationId", op.Text("operationId")
        row.SetText "sheetName", ws.Name
        row.SetText "method", op.Text("method")
        row.SetText "path", op.Text("path")
        row.SetText "summary", op.Text("summary")
        row.SetText "status", op.Text("deprecated")
        indexRows.Push row
    Next i

    mStep = "removed sheets"
    ' Sheets whose operation disappeared from the YAML are kept (they may
    ' hold hand-written content) and marked as removed.
    For Each ws In ThisWorkbook.Worksheets
        key = KeyOf(ws)
        If Left$(key, Len(KEY_PREFIX)) = KEY_PREFIX Then
            If Not names.Has(Mid$(key, Len(KEY_PREFIX) + 1)) Then
                ws.Tab.Color = RGB(160, 160, 160)
                Set row = NewMap()
                row.SetText "no", "-"
                row.SetText "operationId", Mid$(key, Len(KEY_PREFIX) + 1)
                row.SetText "sheetName", ws.Name
                row.SetText "method", ""
                row.SetText "path", ""
                row.SetText "summary", ""
                row.SetText "status", Label("removed")
                indexRows.Push row
            End If
        End If
    Next ws

    ' Index sheet
    mStep = "index sheet"
    model.SetItem "indexRows", indexRows
    Set old = FindSheetByKey(INDEX_KEY)
    Set indexWs = CopyTemplate(TPL_INDEX)
    FillSheet indexWs, NewCtx(model)
    If Not old Is Nothing Then old.Delete
    indexWs.Name = UniqueSheetName(Label("indexSheetName"))
    Set first = FirstVisibleSheet()
    If Not first Is indexWs Then indexWs.Move After:=first

    ' Order: index, then operations in YAML order, then removed ones.
    mStep = "sheet order"
    Set prev = indexWs
    For i = 1 To ops.Count
        Set ws = ThisWorkbook.Worksheets(names.Text(ops.ItemAt(i).Text("operationId")))
        ws.Move After:=prev
        Set prev = ws
    Next i

    indexWs.Activate
    indexWs.Range("A1").Select
    WriteSheets = ops.Count
End Function

Private Function CopyTemplate(ByVal templateName As String) As Worksheet
    Dim tpl As Worksheet, ws As Worksheet, before As clsNode
    ' Do not rely on positions: where Excel puts a copy of a hidden sheet
    ' differs between versions. Find the sheet that was not there before.
    Set before = NewMap()
    For Each ws In ThisWorkbook.Worksheets
        before.SetText ws.Name, ""
    Next ws
    Set tpl = ThisWorkbook.Worksheets(templateName)
    tpl.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    For Each ws In ThisWorkbook.Worksheets
        If Not before.Has(ws.Name) Then
            ws.Visible = xlSheetVisible
            Set CopyTemplate = ws
            Exit Function
        End If
    Next ws
    Err.Raise vbObjectError + 800, "Import", "Could not copy the template sheet " & templateName
End Function

Private Function FirstVisibleSheet() As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            Set FirstVisibleSheet = ws
            Exit Function
        End If
    Next ws
End Function

' The key is in the hidden cell A1: "opid:<operationId>" or "@index".
Private Function KeyOf(ByVal ws As Worksheet) As String
    If ws.Name = TPL_ENDPOINT Or ws.Name = TPL_INDEX Or ws.Name = CONFIG_SHEET Then Exit Function
    KeyOf = CStr(ws.Range("A1").Value)
End Function

Private Function FindSheetByKey(ByVal key As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If KeyOf(ws) = key Then
            Set FindSheetByKey = ws
            Exit Function
        End If
    Next ws
End Function

Private Function UniqueSheetName(ByVal base As String) As String
    Dim bad As Variant, i As Long, candidate As String, n As Long
    bad = Array(":", ChrW(92), "/", "?", "*", "[", "]")
    For i = LBound(bad) To UBound(bad)
        base = Replace(base, bad(i), "_")
    Next i
    If base = "" Then base = "api"
    base = Left$(base, 31)
    candidate = base
    n = 1
    Do While SheetExists(candidate)
        n = n + 1
        candidate = Left$(base, 31 - Len("_" & n)) & "_" & n
    Loop
    UniqueSheetName = candidate
End Function

Private Function SheetExists(ByVal name As String) As Boolean
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, name, vbTextCompare) = 0 Then
            SheetExists = True
            Exit Function
        End If
    Next ws
End Function

' Move the hand-written rows (between the processing markers) from the old
' sheet to the new one. Nothing is moved when the old area is empty, so the
' new sheet keeps the rows made from x-processing / the template.
Private Sub MoveProcessing(ByVal oldWs As Worksheet, ByVal newWs As Worksheet)
    Dim ob As Long, oe As Long, nb As Long, ne As Long, n As Long, i As Long
    ob = FindMark(oldWs, "@processing:begin")
    oe = FindMark(oldWs, "@processing:end")
    nb = FindMark(newWs, "@processing:begin")
    ne = FindMark(newWs, "@processing:end")
    If ob = 0 Or oe = 0 Or nb = 0 Or ne = 0 Then Exit Sub
    n = oe - ob - 1
    If n <= 0 Then Exit Sub
    If Application.WorksheetFunction.CountA(oldWs.Rows((ob + 1) & ":" & (oe - 1))) = 0 Then Exit Sub

    If ne - nb > 1 Then newWs.Rows((nb + 1) & ":" & (ne - 1)).Delete
    oldWs.Rows((ob + 1) & ":" & (oe - 1)).Copy
    newWs.Rows(nb + 1).Insert Shift:=xlDown
    Application.CutCopyMode = False
    For i = 0 To n - 1
        newWs.Rows(nb + 1 + i).RowHeight = oldWs.Rows(ob + 1 + i).RowHeight
        newWs.Rows(nb + 1 + i).Hidden = oldWs.Rows(ob + 1 + i).Hidden
    Next i
End Sub

Private Sub BackupWorkbook()
    Dim folder As String, base As String, p As Long
    If ThisWorkbook.Path = "" Then Exit Sub
    folder = JoinPath(ThisWorkbook.Path, "backup")
    If Not FolderExists(folder) Then MkDir folder
    base = ThisWorkbook.Name
    p = InStrRev(base, ".")
    If p > 0 Then base = Left$(base, p - 1)
    ThisWorkbook.SaveCopyAs JoinPath(folder, base & "_" & Format$(Now, "yyyymmdd-hhnnss") & ".xlsm")
End Sub

' ---------------------------------------------------------------------------
' Settings and labels (sheet "_config": key in column A, value in column B)
' Japanese text lives on that sheet, not in the code.
' ---------------------------------------------------------------------------

Public Function Label(ByVal key As String) As String
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET)
    For r = 1 To LastUsedRow(ws)
        If CStr(ws.Cells(r, 1).Value) = key Then
            Label = CStr(ws.Cells(r, 2).Value)
            Exit Function
        End If
    Next r
    Label = key
End Function

Public Sub SetConfig(ByVal key As String, ByVal value As String)
    Dim ws As Worksheet, r As Long, last As Long
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET)
    last = LastUsedRow(ws)
    For r = 1 To last
        If CStr(ws.Cells(r, 1).Value) = key Then
            ws.Cells(r, 2).Value = "'" & value
            Exit Sub
        End If
    Next r
    ws.Cells(last + 1, 1).Value = key
    ws.Cells(last + 1, 2).Value = "'" & value
End Sub
