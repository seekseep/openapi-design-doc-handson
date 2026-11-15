Attribute VB_Name = "modYaml"
Option Explicit

' Minimal YAML parser covering what OpenAPI documents normally use.
'
' Supported:
'   block maps and sequences, "- key: value" items, sequences at the same
'   indent as their parent key, plain / 'single' / "double" quoted scalars,
'   | and > block scalars (with - / + chomping), flow sequences and maps
'   ([a, b], {a: 1, b: [x]}), comments, a leading "---".
' Not supported (an error with the line number is raised):
'   anchors / aliases, tags, multiple documents, multi-line plain or quoted
'   scalars, complex keys.
'
' All scalars are kept as text. Nulls (an empty value) become Kind = 0.

Private Const ERR_YAML As Long = 513

Private mText() As String     ' content without indentation and comments
Private mRaw() As String      ' original line
Private mIndent() As Long
Private mCount As Long
Private mPos As Long

' state for inline (flow / quoted) parsing
Private mS As String
Private mI As Long
Private mLine As Long

' ---------------------------------------------------------------------------
' Node factories (also used by other modules)
' ---------------------------------------------------------------------------

Public Function NewNull(Optional ByVal srcLine As Long = 0) As clsNode
    Dim n As New clsNode
    n.Kind = 0
    n.SrcLine = srcLine
    Set NewNull = n
End Function

Public Function NewScalar(ByVal v As String, Optional ByVal srcLine As Long = 0) As clsNode
    Dim n As New clsNode
    n.Kind = 1
    n.Value = v
    n.SrcLine = srcLine
    Set NewScalar = n
End Function

Public Function NewMap(Optional ByVal srcLine As Long = 0) As clsNode
    Dim n As New clsNode
    n.Kind = 2
    n.SrcLine = srcLine
    Set NewMap = n
End Function

Public Function NewSeq(Optional ByVal srcLine As Long = 0) As clsNode
    Dim n As New clsNode
    n.Kind = 3
    n.SrcLine = srcLine
    Set NewSeq = n
End Function

' ---------------------------------------------------------------------------
' Entry point
' ---------------------------------------------------------------------------

Public Function YamlParse(ByVal src As String) As clsNode
    SplitLines src
    mPos = 1
    SkipBlank
    If mPos > mCount Then
        Set YamlParse = NewNull(0)
        Exit Function
    End If
    If mText(mPos) = "---" Then
        mPos = mPos + 1
        SkipBlank
    End If
    If mPos > mCount Then
        Set YamlParse = NewNull(0)
        Exit Function
    End If
    Set YamlParse = ParseBlock(mIndent(mPos))
    SkipBlank
    If mPos <= mCount Then
        If mText(mPos) = "---" Then
            Fail mPos, "multiple documents are not supported"
        ElseIf mText(mPos) <> "..." Then
            Fail mPos, "unexpected content (check the indentation)"
        End If
    End If
End Function

Private Sub Fail(ByVal lineNo As Long, ByVal msg As String)
    Err.Raise vbObjectError + ERR_YAML, "YAML", "YAML line " & lineNo & ": " & msg
End Sub

' ---------------------------------------------------------------------------
' Lines
' ---------------------------------------------------------------------------

Private Sub SplitLines(ByVal src As String)
    Dim parts() As String
    Dim i As Long, n As Long, ind As Long
    Dim raw As String, body As String

    src = Replace(src, vbCrLf, vbLf)
    src = Replace(src, vbCr, vbLf)
    parts = Split(src, vbLf)
    n = UBound(parts) - LBound(parts) + 1
    mCount = n
    ReDim mText(1 To n + 1)
    ReDim mRaw(1 To n + 1)
    ReDim mIndent(1 To n + 1)
    For i = 1 To n
        raw = parts(LBound(parts) + i - 1)
        mRaw(i) = raw
        ind = 0
        Do While ind < Len(raw)
            If Mid$(raw, ind + 1, 1) <> " " Then Exit Do
            ind = ind + 1
        Loop
        mIndent(i) = ind
        body = StripComment(Mid$(raw, ind + 1))
        mText(i) = RTrimWs(body)
        If Left$(mText(i), 1) = vbTab Then Fail i, "tabs cannot be used for indentation"
    Next i
End Sub

Private Function RTrimWs(ByVal s As String) As String
    Dim n As Long
    n = Len(s)
    Do While n > 0
        Select Case Mid$(s, n, 1)
            Case " ", vbTab
                n = n - 1
            Case Else
                Exit Do
        End Select
    Loop
    RTrimWs = Left$(s, n)
End Function

' Remove a trailing comment. A quote only opens a quoted scalar at the
' start of a token, so apostrophes inside plain text are not quotes.
Private Function StripComment(ByVal s As String) As String
    Dim i As Long, n As Long
    Dim ch As String, prev As String
    Dim q As String

    n = Len(s)
    i = 1
    prev = " "
    Do While i <= n
        ch = Mid$(s, i, 1)
        If q = "'" Then
            If ch = "'" Then
                If Mid$(s, i + 1, 1) = "'" Then
                    i = i + 1
                Else
                    q = ""
                End If
            End If
        ElseIf q = """" Then
            If ch = BS Then
                i = i + 1
            ElseIf ch = """" Then
                q = ""
            End If
        Else
            If ch = "#" And (prev = " " Or prev = vbTab) Then
                StripComment = Left$(s, i - 1)
                Exit Function
            End If
            If (ch = "'" Or ch = """") And InStr(1, " " & vbTab & ":[{,-", prev, vbBinaryCompare) > 0 Then
                q = ch
            End If
        End If
        prev = ch
        i = i + 1
    Loop
    StripComment = s
End Function

Private Sub SkipBlank()
    Do While mPos <= mCount
        If mText(mPos) <> "" Then Exit Do
        mPos = mPos + 1
    Loop
End Sub

Private Function IsSeqItem(ByVal t As String) As Boolean
    IsSeqItem = (t = "-" Or Left$(t, 2) = "- ")
End Function

' Position of the ":" that separates a key from its value, or 0.
Private Function FindMapColon(ByVal t As String) As Long
    Dim i As Long, n As Long, ch As String, q As String

    n = Len(t)
    If n = 0 Then Exit Function
    ch = Left$(t, 1)
    If ch = "[" Or ch = "{" Then Exit Function
    i = 1
    If ch = "'" Or ch = """" Then
        q = ch
        i = 2
        Do While i <= n
            ch = Mid$(t, i, 1)
            If q = """" And ch = BS Then
                i = i + 1
            ElseIf ch = q Then
                If q = "'" And Mid$(t, i + 1, 1) = "'" Then
                    i = i + 1
                Else
                    Exit Do
                End If
            End If
            i = i + 1
        Loop
        i = i + 1
    End If
    Do While i <= n
        If Mid$(t, i, 1) = ":" Then
            If i = n Then
                FindMapColon = i
                Exit Function
            End If
            ch = Mid$(t, i + 1, 1)
            If ch = " " Or ch = vbTab Then
                FindMapColon = i
                Exit Function
            End If
        End If
        i = i + 1
    Loop
End Function

' ---------------------------------------------------------------------------
' Block structure
' ---------------------------------------------------------------------------

Private Function ParseBlock(ByVal indent As Long) As clsNode
    SkipBlank
    If IsSeqItem(mText(mPos)) Then
        Set ParseBlock = ParseSeq(indent)
    ElseIf FindMapColon(mText(mPos)) > 0 Then
        Set ParseBlock = ParseMap(indent)
    Else
        ' a lone scalar or flow collection as the whole block
        Set ParseBlock = ParseValue(mText(mPos), indent - 1)
    End If
End Function

Private Function ParseSeq(ByVal indent As Long) As clsNode
    Dim seq As clsNode, rest As String, lead As Long, startLine As Long
    Set seq = NewSeq(mPos)
    Do
        SkipBlank
        If mPos > mCount Then Exit Do
        If mIndent(mPos) < indent Then Exit Do
        If mIndent(mPos) > indent Then Fail mPos, "bad indentation"
        If Not IsSeqItem(mText(mPos)) Then Exit Do

        rest = Mid$(mText(mPos), 2)
        lead = Len(rest) - Len(LTrim$(rest))
        rest = LTrim$(rest)
        If rest = "" Then
            startLine = mPos
            mPos = mPos + 1
            SkipBlank
            If mPos > mCount Then
                seq.Push NewNull(startLine)
            ElseIf mIndent(mPos) > indent Then
                seq.Push ParseBlock(mIndent(mPos))
            Else
                seq.Push NewNull(startLine)
            End If
        ElseIf IsSeqItem(rest) Or FindMapColon(rest) > 0 Then
            ' "- key: value" or "- - x": continue as a block that starts here
            mIndent(mPos) = indent + 1 + lead
            mText(mPos) = rest
            seq.Push ParseBlock(mIndent(mPos))
        Else
            seq.Push ParseValue(rest, indent)
        End If
    Loop
    Set ParseSeq = seq
End Function

Private Function ParseMap(ByVal indent As Long) As clsNode
    Dim m As clsNode, v As clsNode
    Dim t As String, key As String, rest As String
    Dim c As Long, startLine As Long

    Set m = NewMap(mPos)
    Do
        SkipBlank
        If mPos > mCount Then Exit Do
        If mIndent(mPos) < indent Then Exit Do
        If mIndent(mPos) > indent Then Fail mPos, "bad indentation"
        t = mText(mPos)
        If IsSeqItem(t) Then Exit Do
        ' document markers end the mapping; YamlParse reports what follows
        If t = "---" Or t = "..." Then Exit Do
        c = FindMapColon(t)
        If c = 0 Then Fail mPos, "expected 'key: value' but got '" & t & "'"
        startLine = mPos
        key = ParseKey(RTrimWs(Left$(t, c - 1)), mPos)
        rest = Trim$(Mid$(t, c + 1))
        If Left$(rest, 1) = "&" Or Left$(rest, 1) = "*" Then Fail mPos, "anchors and aliases are not supported"
        If Left$(rest, 1) = "!" Then Fail mPos, "tags are not supported"

        If rest = "" Then
            mPos = mPos + 1
            SkipBlank
            If mPos > mCount Then
                Set v = NewNull(startLine)
            ElseIf mIndent(mPos) > indent Then
                Set v = ParseBlock(mIndent(mPos))
            ElseIf mIndent(mPos) = indent And IsSeqItem(mText(mPos)) Then
                Set v = ParseSeq(indent)
            Else
                Set v = NewNull(startLine)
            End If
        Else
            Set v = ParseValue(rest, indent)
        End If
        If m.Has(key) Then Fail startLine, "duplicate key '" & key & "'"
        m.SetItem key, v
    Loop
    Set ParseMap = m
End Function

Private Function ParseKey(ByVal k As String, ByVal lineNo As Long) As String
    Dim ch As String
    ch = Left$(k, 1)
    If ch = "'" Or ch = """" Then
        ParseKey = ParseInline(k, lineNo).Value
    Else
        If ch = "?" Then Fail lineNo, "complex keys are not supported"
        ParseKey = k
    End If
End Function

' Value written after "key:" or "- ". Advances mPos past the value.
Private Function ParseValue(ByVal rest As String, ByVal parentIndent As Long) As clsNode
    Dim ch As String
    ch = Left$(rest, 1)
    If (ch = "|" Or ch = ">") And IsBlockHeader(rest) Then
        Set ParseValue = ParseBlockScalar(rest, parentIndent)
    Else
        Set ParseValue = ParseInline(rest, mPos)
        mPos = mPos + 1
    End If
End Function

Private Function IsBlockHeader(ByVal h As String) As Boolean
    Dim i As Long, ch As String
    For i = 2 To Len(h)
        ch = Mid$(h, i, 1)
        If InStr(1, "+-0123456789", ch, vbBinaryCompare) = 0 Then
            IsBlockHeader = False
            Exit Function
        End If
    Next i
    IsBlockHeader = True
End Function

Private Function ParseBlockScalar(ByVal header As String, ByVal parentIndent As Long) As clsNode
    Dim startLine As Long, blockIndent As Long, ind As Long
    Dim raw As String, folded As Boolean, chomp As String
    Dim lines() As String, n As Long, i As Long
    Dim result As String, prevBlank As Boolean

    startLine = mPos
    folded = (Left$(header, 1) = ">")
    If InStr(header, "-") > 0 Then
        chomp = "-"
    ElseIf InStr(header, "+") > 0 Then
        chomp = "+"
    End If

    ReDim lines(1 To 16)
    mPos = mPos + 1
    Do While mPos <= mCount
        raw = mRaw(mPos)
        If Trim$(raw) = "" Then
            n = n + 1
            If n > UBound(lines) Then ReDim Preserve lines(1 To UBound(lines) * 2)
            lines(n) = ""
        Else
            ind = 0
            Do While ind < Len(raw)
                If Mid$(raw, ind + 1, 1) <> " " Then Exit Do
                ind = ind + 1
            Loop
            If blockIndent = 0 Then
                If ind <= parentIndent Then Exit Do
                blockIndent = ind
            End If
            If ind < blockIndent Then Exit Do
            n = n + 1
            If n > UBound(lines) Then ReDim Preserve lines(1 To UBound(lines) * 2)
            lines(n) = Mid$(raw, blockIndent + 1)
        End If
        mPos = mPos + 1
    Loop

    ' Trailing blank lines are not part of the content.
    Dim trailing As Long
    Do While n > 0
        If lines(n) <> "" Then Exit Do
        n = n - 1
        trailing = trailing + 1
    Loop
    ' Give the consumed trailing blank lines back to the structure parser.
    mPos = mPos - trailing

    For i = 1 To n
        If Not folded Then
            If i > 1 Then result = result & vbLf
            result = result & lines(i)
        Else
            If lines(i) = "" Then
                result = result & vbLf
                prevBlank = True
            Else
                If i > 1 And Not prevBlank Then result = result & " "
                result = result & lines(i)
                prevBlank = False
            End If
        End If
    Next i

    If n > 0 Then
        If chomp = "" Then
            result = result & vbLf
        ElseIf chomp = "+" Then
            result = result & vbLf & String$(trailing, vbLf)
        End If
    End If
    Set ParseBlockScalar = NewScalar(result, startLine)
End Function

' ---------------------------------------------------------------------------
' Inline values: plain, quoted and flow collections
' ---------------------------------------------------------------------------

Private Function ParseInline(ByVal s As String, ByVal lineNo As Long) As clsNode
    Dim n As clsNode
    mS = s
    mI = 1
    mLine = lineNo
    Set n = FlowValue(False)
    SkipWs
    If mI <= Len(mS) Then Fail mLine, "unexpected text '" & Mid$(mS, mI) & "'"
    Set ParseInline = n
End Function

Private Sub SkipWs()
    Do While mI <= Len(mS)
        Select Case Mid$(mS, mI, 1)
            Case " ", vbTab
                mI = mI + 1
            Case Else
                Exit Do
        End Select
    Loop
End Sub

Private Function FlowValue(ByVal inFlow As Boolean) As clsNode
    Dim ch As String
    SkipWs
    ch = Mid$(mS, mI, 1)
    Select Case ch
        Case "["
            Set FlowValue = FlowSeq()
        Case "{"
            Set FlowValue = FlowMap()
        Case "'"
            Set FlowValue = NewScalar(SingleQuoted(), mLine)
        Case """"
            Set FlowValue = NewScalar(DoubleQuoted(), mLine)
        Case "&", "*"
            Fail mLine, "anchors and aliases are not supported"
        Case Else
            Set FlowValue = PlainScalar(inFlow, False)
    End Select
End Function

Private Function PlainScalar(ByVal inFlow As Boolean, ByVal isKey As Boolean) As clsNode
    Dim start As Long, ch As String, v As String
    start = mI
    Do While mI <= Len(mS)
        ch = Mid$(mS, mI, 1)
        If inFlow Then
            If ch = "," Or ch = "]" Or ch = "}" Then Exit Do
            If isKey And ch = ":" Then Exit Do
            If ch = ":" And (Mid$(mS, mI + 1, 1) = " " Or Mid$(mS, mI + 1, 1) = "") And Not isKey Then Exit Do
        End If
        mI = mI + 1
    Loop
    v = Trim$(Mid$(mS, start, mI - start))
    If v = "" Or v = "~" Or v = "null" Or v = "Null" Or v = "NULL" Then
        If v = "" Then
            Set PlainScalar = NewNull(mLine)
        Else
            Set PlainScalar = NewScalar(v, mLine)
        End If
    Else
        Set PlainScalar = NewScalar(v, mLine)
    End If
End Function

Private Function FlowSeq() As clsNode
    Dim seq As clsNode, ch As String
    Set seq = NewSeq(mLine)
    mI = mI + 1
    Do
        SkipWs
        If mI > Len(mS) Then Fail mLine, "unterminated '['"
        ch = Mid$(mS, mI, 1)
        If ch = "]" Then
            mI = mI + 1
            Exit Do
        End If
        seq.Push FlowValue(True)
        SkipWs
        ch = Mid$(mS, mI, 1)
        If ch = "," Then
            mI = mI + 1
        ElseIf ch = "]" Then
            mI = mI + 1
            Exit Do
        Else
            Fail mLine, "expected ',' or ']' in flow sequence"
        End If
    Loop
    Set FlowSeq = seq
End Function

Private Function FlowMap() As clsNode
    Dim m As clsNode, ch As String, key As String, v As clsNode
    Set m = NewMap(mLine)
    mI = mI + 1
    Do
        SkipWs
        If mI > Len(mS) Then Fail mLine, "unterminated '{'"
        ch = Mid$(mS, mI, 1)
        If ch = "}" Then
            mI = mI + 1
            Exit Do
        End If
        If ch = "'" Then
            key = SingleQuoted()
        ElseIf ch = """" Then
            key = DoubleQuoted()
        Else
            key = PlainScalar(True, True).Value
        End If
        SkipWs
        If Mid$(mS, mI, 1) = ":" Then
            mI = mI + 1
            SkipWs
            ch = Mid$(mS, mI, 1)
            If ch = "," Or ch = "}" Then
                Set v = NewNull(mLine)
            Else
                Set v = FlowValue(True)
            End If
        Else
            Set v = NewNull(mLine)
        End If
        If m.Has(key) Then Fail mLine, "duplicate key '" & key & "'"
        m.SetItem key, v
        SkipWs
        ch = Mid$(mS, mI, 1)
        If ch = "," Then
            mI = mI + 1
        ElseIf ch = "}" Then
            mI = mI + 1
            Exit Do
        Else
            Fail mLine, "expected ',' or '}' in flow map"
        End If
    Loop
    Set FlowMap = m
End Function

Private Function SingleQuoted() As String
    Dim r As String, ch As String
    mI = mI + 1
    Do
        If mI > Len(mS) Then Fail mLine, "unterminated quoted string"
        ch = Mid$(mS, mI, 1)
        If ch = "'" Then
            If Mid$(mS, mI + 1, 1) = "'" Then
                r = r & "'"
                mI = mI + 2
            Else
                mI = mI + 1
                Exit Do
            End If
        Else
            r = r & ch
            mI = mI + 1
        End If
    Loop
    SingleQuoted = r
End Function

Private Function DoubleQuoted() As String
    Dim r As String, ch As String, e As String
    mI = mI + 1
    Do
        If mI > Len(mS) Then Fail mLine, "unterminated quoted string"
        ch = Mid$(mS, mI, 1)
        If ch = """" Then
            mI = mI + 1
            Exit Do
        ElseIf ch = BS Then
            e = Mid$(mS, mI + 1, 1)
            Select Case e
                Case "n": r = r & vbLf
                Case "t": r = r & vbTab
                Case "r": r = r & vbCr
                Case "0": r = r & ChrW(0)
                Case """", BS, "/", " ": r = r & e
                Case "u"
                    r = r & ChrW(CLng("&H" & Mid$(mS, mI + 2, 4)))
                    mI = mI + 4
                Case Else
                    Fail mLine, "unsupported escape " & BS & e
            End Select
            mI = mI + 2
        Else
            r = r & ch
            mI = mI + 1
        End If
    Loop
    DoubleQuoted = r
End Function

' ---------------------------------------------------------------------------
' Canonical dump (used by tests to compare with another YAML parser)
' ---------------------------------------------------------------------------

Public Function YamlDump(ByVal n As clsNode) As String
    Dim i As Long, s As String
    Select Case n.Kind
        Case 0
            YamlDump = """"""
        Case 1
            YamlDump = DumpString(n.Value)
        Case 2
            s = "{"
            For i = 1 To n.Count
                If i > 1 Then s = s & ","
                s = s & DumpString(n.KeyAt(i)) & ":" & YamlDump(n.ItemAt(i))
            Next i
            YamlDump = s & "}"
        Case 3
            s = "["
            For i = 1 To n.Count
                If i > 1 Then s = s & ","
                s = s & YamlDump(n.ItemAt(i))
            Next i
            YamlDump = s & "]"
    End Select
End Function

Private Function DumpString(ByVal v As String) As String
    v = Replace(v, BS, BS & BS)
    v = Replace(v, """", BS & """")
    v = Replace(v, vbLf, BS & "n")
    DumpString = """" & v & """"
End Function

' The backslash is written as ChrW(92): on Japanese Mac Excel the byte 0x5C
' in source code is read as a yen sign, so a literal backslash breaks.
Private Function BS() As String
    BS = ChrW(92)
End Function
