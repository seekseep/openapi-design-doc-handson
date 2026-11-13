Attribute VB_Name = "modTemplate"
Option Explicit

' Fill a sheet that was copied from a template sheet.
'
' Column A (hidden) holds a marker for each row:
'   @row:<list>        repeat this row for each item of <list>
'   @empty:<list>      keep this row only when <list> is empty
'   @if:<key>          keep this row only when <key> has a value
'   @block:<list> ... @end:<list>    repeat the rows in between for each item
'   @section:<list> ... @end:<list>  keep the rows in between once when <list> is not empty
'   anything else      a normal row (markers such as @processing:begin stay as they are)
' Cells may contain {{key}}. The value is looked up from the innermost
' context outwards. {{sheetLink}} becomes a link to the sheet named by
' the "sheetName" value.

Private Const MARK_COL As Long = 1

Private mLastCol As Long

Public Sub FillSheet(ByVal ws As Worksheet, ByVal ctx As Collection)
    mLastCol = LastUsedCol(ws)
    ExpandRange ws, 1, LastUsedRow(ws), ctx
    Application.CutCopyMode = False
End Sub

Public Function LastUsedRow(ByVal ws As Worksheet) As Long
    Dim r As Range
    Set r = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If r Is Nothing Then
        LastUsedRow = 1
    Else
        LastUsedRow = r.Row
    End If
End Function

Private Function LastUsedCol(ByVal ws As Worksheet) As Long
    Dim r As Range
    Set r = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If r Is Nothing Then
        LastUsedCol = 1
    Else
        LastUsedCol = r.Column
    End If
End Function

' Row number of the first row whose marker is exactly <mark>, or 0.
Public Function FindMark(ByVal ws As Worksheet, ByVal mark As String) As Long
    Dim r As Long, last As Long
    last = LastUsedRow(ws)
    For r = 1 To last
        If CStr(ws.Cells(r, MARK_COL).Value) = mark Then
            FindMark = r
            Exit Function
        End If
    Next r
End Function

' Returns the last row of the range after expansion.
Private Function ExpandRange(ByVal ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long, _
                             ByVal ctx As Collection) As Long
    Dim r As Long, kind As String, name As String
    Dim list As clsNode, n As Long, k As Long, e As Long, blockLen As Long
    Dim startRow As Long, newEnd As Long

    r = r1
    Do While r <= r2
        SplitMark CStr(ws.Cells(r, MARK_COL).Value), kind, name
        Select Case kind
            Case "row"
                Set list = LookupNode(ctx, name)
                n = ListCount(list)
                If n = 0 Then
                    ws.Rows(r).Delete
                    r2 = r2 - 1
                Else
                    For k = 1 To n - 1
                        ws.Rows(r).Copy
                        ws.Rows(r + 1).Insert Shift:=xlDown
                    Next k
                    For k = 1 To n
                        ClearMark ws, r + k - 1
                        FillRow ws, r + k - 1, PushCtx(ctx, list.ItemAt(k))
                    Next k
                    r = r + n
                    r2 = r2 + n - 1
                End If

            Case "empty", "if"
                Set list = LookupNode(ctx, name)
                If (kind = "empty" And ListCount(list) = 0) Or (kind = "if" And ListCount(list) > 0) Then
                    ClearMark ws, r
                    FillRow ws, r, ctx
                    r = r + 1
                Else
                    ws.Rows(r).Delete
                    r2 = r2 - 1
                End If

            Case "block", "section"
                e = FindEnd(ws, r, r2, name)
                If e = 0 Then Err.Raise vbObjectError + 700, "Template", "@end:" & name & " not found (row " & r & ")"
                Set list = LookupNode(ctx, name)
                If kind = "block" Then
                    n = ListCount(list)
                ElseIf ListCount(list) > 0 Then
                    n = 1
                Else
                    n = 0
                End If
                ws.Rows(e).Delete
                ws.Rows(r).Delete
                r2 = r2 - 2
                e = e - 2
                blockLen = e - r + 1
                If n = 0 Then
                    If blockLen > 0 Then ws.Rows(r & ":" & e).Delete
                    r2 = r2 - blockLen
                ElseIf blockLen > 0 Then
                    For k = 1 To n - 1
                        ws.Rows(r & ":" & e).Copy
                        ws.Rows(e + 1).Insert Shift:=xlDown
                    Next k
                    r2 = r2 + blockLen * (n - 1)
                    startRow = r
                    For k = 1 To n
                        If kind = "block" Then
                            newEnd = ExpandRange(ws, startRow, startRow + blockLen - 1, PushCtx(ctx, list.ItemAt(k)))
                        Else
                            newEnd = ExpandRange(ws, startRow, startRow + blockLen - 1, ctx)
                        End If
                        r2 = r2 + (newEnd - (startRow + blockLen - 1))
                        startRow = newEnd + 1
                    Next k
                    r = startRow
                End If

            Case Else
                FillRow ws, r, ctx
                r = r + 1
        End Select
    Loop
    ExpandRange = r2
End Function

Private Sub SplitMark(ByVal mark As String, ByRef kind As String, ByRef name As String)
    Dim p As Long
    kind = ""
    name = ""
    If Left$(mark, 1) <> "@" Then Exit Sub
    p = InStr(mark, ":")
    If p = 0 Then Exit Sub
    kind = Mid$(mark, 2, p - 2)
    name = Mid$(mark, p + 1)
    Select Case kind
        Case "row", "empty", "if", "block", "section", "end"
        Case Else
            kind = ""
    End Select
End Sub

Private Function FindEnd(ByVal ws As Worksheet, ByVal r As Long, ByVal r2 As Long, ByVal name As String) As Long
    Dim i As Long, depth As Long, m As String
    For i = r + 1 To r2
        m = CStr(ws.Cells(i, MARK_COL).Value)
        If m = "@block:" & name Or m = "@section:" & name Then
            depth = depth + 1
        ElseIf m = "@end:" & name Then
            If depth = 0 Then
                FindEnd = i
                Exit Function
            End If
            depth = depth - 1
        End If
    Next i
End Function

Private Sub ClearMark(ByVal ws As Worksheet, ByVal r As Long)
    ws.Cells(r, MARK_COL).ClearContents
    ws.Rows(r).Hidden = False
End Sub

' ---------------------------------------------------------------------------
' Context (a stack of clsNode maps, innermost last)
' ---------------------------------------------------------------------------

Public Function NewCtx(ParamArray nodes() As Variant) As Collection
    Dim c As New Collection, i As Long
    For i = LBound(nodes) To UBound(nodes)
        c.Add nodes(i)
    Next i
    Set NewCtx = c
End Function

Private Function PushCtx(ByVal ctx As Collection, ByVal node As clsNode) As Collection
    Dim c As New Collection, i As Long
    For i = 1 To ctx.Count
        c.Add ctx(i)
    Next i
    c.Add node
    Set PushCtx = c
End Function

Private Function LookupNode(ByVal ctx As Collection, ByVal key As String) As clsNode
    Dim i As Long, n As clsNode
    For i = ctx.Count To 1 Step -1
        Set n = ctx(i)
        If n.IsMap Then
            If n.Has(key) Then
                Set LookupNode = n.Child(key)
                Exit Function
            End If
        End If
    Next i
    Set LookupNode = Nothing
End Function

Private Function LookupText(ByVal ctx As Collection, ByVal key As String) As String
    Dim n As clsNode
    Set n = LookupNode(ctx, key)
    If n Is Nothing Then Exit Function
    If n.IsScalar Then LookupText = n.Value
End Function

' Sequences: item count. Scalars: 1 when not empty.
Private Function ListCount(ByVal n As clsNode) As Long
    If n Is Nothing Then Exit Function
    If n.IsSeq Or n.IsMap Then
        ListCount = n.Count
    ElseIf n.IsScalar Then
        If n.Value <> "" Then ListCount = 1
    End If
End Function

' ---------------------------------------------------------------------------
' Cells
' ---------------------------------------------------------------------------

Private Sub FillRow(ByVal ws As Worksheet, ByVal r As Long, ByVal ctx As Collection)
    Dim c As Long, cell As Range, v As Variant, s As String, changed As Boolean, sheetName As String

    For c = 1 To mLastCol
        Set cell = ws.Cells(r, c)
        If Not cell.HasFormula Then
            v = cell.Value
            If VarType(v) = vbString Then
                s = v
                If InStr(s, "{{") > 0 Then
                    If s = "{{sheetLink}}" Then
                        sheetName = LookupText(ctx, "sheetName")
                        cell.Value = "'" & LookupText(ctx, "operationId")
                        If sheetName <> "" Then
                            ws.Hyperlinks.Add Anchor:=cell, Address:="", _
                                SubAddress:="'" & sheetName & "'!A1", TextToDisplay:=LookupText(ctx, "operationId")
                        End If
                    Else
                        s = ReplacePlaceholders(s, ctx)
                        If s = "" Then
                            cell.ClearContents
                        Else
                            cell.Value = "'" & s
                        End If
                    End If
                    changed = True
                End If
            End If
        End If
    Next c
    If changed And Not ws.Rows(r).Hidden Then FitRowHeight ws, r
End Sub

Private Function ReplacePlaceholders(ByVal s As String, ByVal ctx As Collection) As String
    Dim p As Long, q As Long, key As String, result As String
    Do
        p = InStr(s, "{{")
        If p = 0 Then Exit Do
        q = InStr(p + 2, s, "}}")
        If q = 0 Then Exit Do
        key = Trim$(Mid$(s, p + 2, q - p - 2))
        result = result & Left$(s, p - 1) & LookupText(ctx, key)
        s = Mid$(s, q + 2)
    Loop
    ReplacePlaceholders = result & s
End Function

' AutoFit ignores merged cells, so estimate their height from the text.
Private Sub FitRowHeight(ByVal ws As Worksheet, ByVal r As Long)
    Dim c As Long, cell As Range, area As Range, lines As Long, maxLines As Long
    Dim colWidth As Double, k As Long, baseHeight As Double

    baseHeight = ws.StandardHeight
    ws.Rows(r).AutoFit
    maxLines = 1
    For c = 1 To mLastCol
        Set cell = ws.Cells(r, c)
        If cell.MergeCells Then
            Set area = cell.MergeArea
            If area.Cells(1, 1).Address = cell.Address And cell.WrapText Then
                colWidth = 0
                For k = 1 To area.Columns.Count
                    colWidth = colWidth + area.Columns(k).ColumnWidth
                Next k
                lines = CountLines(CStr(cell.Value), colWidth)
                If lines > maxLines Then maxLines = lines
            End If
        End If
    Next c
    If maxLines > 1 Then
        If ws.Rows(r).RowHeight < baseHeight * maxLines Then ws.Rows(r).RowHeight = baseHeight * maxLines
    End If
End Sub

Private Function CountLines(ByVal text As String, ByVal colWidth As Double) As Long
    Dim parts() As String, i As Long, units As Double, j As Long, total As Long
    If colWidth < 1 Then colWidth = 1
    parts = Split(text, vbLf)
    For i = LBound(parts) To UBound(parts)
        units = 0
        For j = 1 To Len(parts(i))
            If AscW(Mid$(parts(i), j, 1)) > 255 Or AscW(Mid$(parts(i), j, 1)) < 0 Then
                units = units + 2
            Else
                units = units + 1
            End If
        Next j
        If units = 0 Then
            total = total + 1
        Else
            total = total + Int((units - 1) / colWidth) + 1
        End If
    Next i
    If total < 1 Then total = 1
    CountLines = total
End Function
