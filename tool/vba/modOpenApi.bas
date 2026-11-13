Attribute VB_Name = "modOpenApi"
Option Explicit

' Turn a parsed OpenAPI document into the values the template needs.
'
' The result is a tree of clsNode maps / sequences:
'   model: title, version, sourcePath, importedAt, operations[]
'   operation: operationId, method, path, tags, summary, purpose, description,
'              deprecated, auth, permissions[], params[], bodies[],
'              responses[], processing[]
'   rows (params / fields / headers): name, in, type, required, constraints,
'              description, example
' Every list key is always present (maybe empty) so that template lookups
' never fall through to an outer context by accident.

Private Const HTTP_METHODS As String = "|get|put|post|delete|options|head|patch|trace|"
Private Const MAX_DEPTH As Long = 8

Private mRoot As clsNode

Public Function BuildModel(ByVal root As clsNode, ByVal sourcePath As String) As clsNode
    Dim model As clsNode, ops As clsNode, paths As clsNode, info As clsNode
    Dim pathItem As clsNode, i As Long, j As Long, method As String

    If Not root.IsMap Then Err.Raise vbObjectError + 600, "OpenAPI", "The file is not an OpenAPI document."
    Set mRoot = root
    Set model = NewMap()
    Set info = root.Child("info")
    If info Is Nothing Then
        model.SetText "title", ""
        model.SetText "version", ""
    Else
        model.SetText "title", info.Text("title")
        model.SetText "version", info.Text("version")
    End If
    model.SetText "sourcePath", sourcePath
    model.SetText "importedAt", Format$(Now, "yyyy-mm-dd hh:nn:ss")

    Set ops = NewSeq()
    Set paths = root.Child("paths")
    If paths Is Nothing Then Err.Raise vbObjectError + 601, "OpenAPI", "'paths' was not found."
    For i = 1 To paths.Count
        Set pathItem = Resolve(paths.ItemAt(i))
        If pathItem.IsMap Then
            For j = 1 To pathItem.Count
                method = LCase$(pathItem.KeyAt(j))
                If InStr(1, HTTP_METHODS, "|" & method & "|", vbBinaryCompare) > 0 Then
                    ops.Push BuildOperation(paths.KeyAt(i), method, pathItem, Resolve(pathItem.ItemAt(j)))
                End If
            Next j
        End If
    Next i
    model.SetItem "operations", ops
    Set BuildModel = model
End Function

Private Function BuildOperation(ByVal pathKey As String, ByVal method As String, _
                                ByVal pathItem As clsNode, ByVal op As clsNode) As clsNode
    Dim o As clsNode, opId As String, purpose As String

    Set o = NewMap()
    opId = op.Text("operationId")
    If opId = "" Then opId = method & CleanId(pathKey)
    o.SetText "operationId", opId
    o.SetText "method", UCase$(method)
    o.SetText "path", pathKey
    o.SetText "tags", JoinScalars(op.Child("tags"), ", ")
    o.SetText "summary", op.Text("summary")
    purpose = op.Text("x-purpose")
    If purpose = "" Then purpose = op.Text("description")
    o.SetText "purpose", TrimText(purpose)
    o.SetText "description", TrimText(op.Text("description"))
    If op.Text("deprecated") = "true" Then
        o.SetText "deprecated", Label("deprecated")
    Else
        o.SetText "deprecated", ""
    End If
    o.SetText "auth", AuthText(op)
    o.SetItem "permissions", Permissions(op)
    o.SetItem "params", Params(pathItem, op)
    o.SetItem "bodies", Bodies(op)
    o.SetItem "responses", Responses(op)
    o.SetItem "processing", Processing(op)
    Set BuildOperation = o
End Function

' ---------------------------------------------------------------------------
' $ref
' ---------------------------------------------------------------------------

Public Function RefName(ByVal n As clsNode) As String
    Dim ref As String
    If n Is Nothing Then Exit Function
    If Not n.IsMap Then Exit Function
    ref = n.Text("$ref")
    If ref = "" Then Exit Function
    RefName = Mid$(ref, InStrRev(ref, "/") + 1)
End Function

Public Function Resolve(ByVal n As clsNode) As clsNode
    Dim guard As Long, ref As String
    If n Is Nothing Then
        Set Resolve = NewMap()
        Exit Function
    End If
    Do While n.IsMap
        ref = n.Text("$ref")
        If ref = "" Then Exit Do
        Set n = Lookup(ref)
        guard = guard + 1
        If guard > 20 Then Err.Raise vbObjectError + 602, "OpenAPI", "$ref loop: " & ref
    Loop
    Set Resolve = n
End Function

Private Function Lookup(ByVal ref As String) As clsNode
    Dim parts() As String, i As Long, n As clsNode, key As String
    If Left$(ref, 2) <> "#/" Then
        Err.Raise vbObjectError + 603, "OpenAPI", "Only local $ref (#/...) is supported: " & ref
    End If
    parts = Split(Mid$(ref, 3), "/")
    Set n = mRoot
    For i = LBound(parts) To UBound(parts)
        key = Replace(Replace(parts(i), "~1", "/"), "~0", "~")
        If Not n.IsMap Then Set n = Nothing
        If Not n Is Nothing Then Set n = n.Child(key)
        If n Is Nothing Then Err.Raise vbObjectError + 604, "OpenAPI", "$ref target not found: " & ref
    Next i
    Set Lookup = n
End Function

' ---------------------------------------------------------------------------
' Security / permissions / processing
' ---------------------------------------------------------------------------

Private Function AuthText(ByVal op As clsNode) As String
    Dim sec As clsNode, req As clsNode, i As Long, j As Long
    Dim alt As String, result As String

    Set sec = op.Child("security")
    If sec Is Nothing Then Set sec = mRoot.Child("security")
    If sec Is Nothing Then
        AuthText = Label("authNone")
        Exit Function
    End If
    If Not sec.IsSeq Or sec.Count = 0 Then
        AuthText = Label("authNone")
        Exit Function
    End If
    For i = 1 To sec.Count
        Set req = sec.ItemAt(i)
        alt = ""
        If req.IsMap Then
            For j = 1 To req.Count
                If alt <> "" Then alt = alt & " + "
                alt = alt & SchemeText(req.KeyAt(j), req.ItemAt(j))
            Next j
        End If
        If alt = "" Then alt = Label("authNone")
        If result <> "" Then result = result & vbLf
        result = result & alt
    Next i
    AuthText = result
End Function

Private Function SchemeText(ByVal name As String, ByVal scopes As clsNode) As String
    Dim schemes As clsNode, s As clsNode, t As String, d As String

    Set schemes = mRoot.Child("components")
    If Not schemes Is Nothing Then Set schemes = schemes.Child("securitySchemes")
    If Not schemes Is Nothing Then Set s = schemes.Child(name)
    If s Is Nothing Then
        SchemeText = name
        Exit Function
    End If
    Set s = Resolve(s)
    t = s.Text("type")
    Select Case t
        Case "http"
            d = "HTTP " & s.Text("scheme")
            If s.Text("bearerFormat") <> "" Then d = d & " (" & s.Text("bearerFormat") & ")"
        Case "apiKey"
            d = "API key (" & s.Text("in") & ": " & s.Text("name") & ")"
        Case Else
            d = t
    End Select
    SchemeText = name & ": " & d
    If Not scopes Is Nothing Then
        If scopes.IsSeq And scopes.Count > 0 Then SchemeText = SchemeText & " [" & JoinScalars(scopes, ", ") & "]"
    End If
End Function

Private Function Permissions(ByVal op As clsNode) As clsNode
    Dim list As clsNode, src As clsNode, item As clsNode, row As clsNode, i As Long
    Set list = NewSeq()
    Set src = op.Child("x-permissions")
    If Not src Is Nothing Then
        If src.IsSeq Then
            For i = 1 To src.Count
                Set item = src.ItemAt(i)
                Set row = NewMap()
                If item.IsMap Then
                    row.SetText "role", item.Text("role")
                    row.SetText "condition", TrimText(item.Text("condition"))
                Else
                    row.SetText "role", item.Value
                    row.SetText "condition", ""
                End If
                list.Push row
            Next i
        End If
    End If
    Set Permissions = list
End Function

Private Function Processing(ByVal op As clsNode) As clsNode
    Dim list As clsNode, src As clsNode, item As clsNode, row As clsNode, i As Long
    Set list = NewSeq()
    Set src = op.Child("x-processing")
    If Not src Is Nothing Then
        If src.IsSeq Then
            For i = 1 To src.Count
                Set item = src.ItemAt(i)
                Set row = NewMap()
                row.SetText "no", CStr(i)
                If item.IsMap Then
                    row.SetText "step", item.Text("step")
                    row.SetText "detail", TrimText(item.Text("detail"))
                    row.SetText "note", TrimText(item.Text("note"))
                Else
                    row.SetText "step", item.Value
                    row.SetText "detail", ""
                    row.SetText "note", ""
                End If
                list.Push row
            Next i
        End If
    End If
    Set Processing = list
End Function

' ---------------------------------------------------------------------------
' Parameters / bodies / responses
' ---------------------------------------------------------------------------

Private Function Params(ByVal pathItem As clsNode, ByVal op As clsNode) As clsNode
    Dim merged As clsNode, rows As clsNode, i As Long
    Set merged = NewMap()
    AddParams merged, pathItem.Child("parameters")
    AddParams merged, op.Child("parameters")   ' operation level wins
    Set rows = NewSeq()
    For i = 1 To merged.Count
        rows.Push ParamRow(merged.ItemAt(i))
    Next i
    Set Params = rows
End Function

Private Sub AddParams(ByVal merged As clsNode, ByVal list As clsNode)
    Dim i As Long, p As clsNode
    If list Is Nothing Then Exit Sub
    If Not list.IsSeq Then Exit Sub
    For i = 1 To list.Count
        Set p = Resolve(list.ItemAt(i))
        merged.SetItem p.Text("in") & ":" & p.Text("name"), p
    Next i
End Sub

Private Function ParamRow(ByVal p As clsNode) As clsNode
    Dim row As clsNode, schema As clsNode, c As String, desc As String, ex As String
    Set row = NewMap()
    Set schema = p.Child("schema")
    row.SetText "name", p.Text("name")
    row.SetText "in", p.Text("in")
    row.SetText "type", TypeText(schema)
    If p.Text("required") = "true" Or p.Text("in") = "path" Then
        row.SetText "required", Label("yes")
    Else
        row.SetText "required", ""
    End If
    c = Constraints(schema)
    If p.Text("style") <> "" Then c = AddLine(c, "style: " & p.Text("style"))
    If p.Text("explode") <> "" Then c = AddLine(c, "explode: " & p.Text("explode"))
    If p.Text("deprecated") = "true" Then c = AddLine(c, "deprecated")
    row.SetText "constraints", c
    desc = TrimText(p.Text("description"))
    If desc = "" Then desc = TrimText(Resolve(schema).Text("description"))
    row.SetText "description", desc
    ex = ExampleText(p.Child("example"))
    If ex = "" Then ex = ExampleText(Resolve(schema).Child("example"))
    row.SetText "example", ex
    Set ParamRow = row
End Function

Private Function Bodies(ByVal op As clsNode) As clsNode
    Dim list As clsNode, rb As clsNode, content As clsNode, b As clsNode, i As Long
    Set list = NewSeq()
    If Not op.Child("requestBody") Is Nothing Then
        Set rb = Resolve(op.Child("requestBody"))
        Set content = rb.Child("content")
        If Not content Is Nothing Then
            For i = 1 To content.Count
                Set b = NewMap()
                b.SetText "contentType", content.KeyAt(i)
                If rb.Text("required") = "true" Then
                    b.SetText "required", Label("yes")
                Else
                    b.SetText "required", ""
                End If
                b.SetText "description", TrimText(rb.Text("description"))
                b.SetItem "fields", BodyRows(content.ItemAt(i).Child("schema"))
                list.Push b
            Next i
        End If
    End If
    Set Bodies = list
End Function

Private Function Responses(ByVal op As clsNode) As clsNode
    Dim list As clsNode, src As clsNode, r As clsNode, item As clsNode
    Dim contents As clsNode, content As clsNode, c As clsNode, i As Long, j As Long

    Set list = NewSeq()
    Set src = op.Child("responses")
    If Not src Is Nothing Then
        For i = 1 To src.Count
            Set r = Resolve(src.ItemAt(i))
            Set item = NewMap()
            item.SetText "status", src.KeyAt(i)
            item.SetText "statusDescription", TrimText(r.Text("description"))
            item.SetItem "headers", HeaderRows(r.Child("headers"))
            Set contents = NewSeq()
            Set content = r.Child("content")
            If Not content Is Nothing Then
                For j = 1 To content.Count
                    Set c = NewMap()
                    c.SetText "contentType", content.KeyAt(j)
                    c.SetItem "fields", BodyRows(content.ItemAt(j).Child("schema"))
                    contents.Push c
                Next j
            End If
            item.SetItem "contents", contents
            list.Push item
        Next i
    End If
    Set Responses = list
End Function

Private Function HeaderRows(ByVal headers As clsNode) As clsNode
    Dim rows As clsNode, h As clsNode, row As clsNode, i As Long
    Set rows = NewSeq()
    If Not headers Is Nothing Then
        For i = 1 To headers.Count
            Set h = Resolve(headers.ItemAt(i))
            Set row = NewMap()
            row.SetText "name", headers.KeyAt(i)
            row.SetText "in", "header"
            row.SetText "type", TypeText(h.Child("schema"))
            If h.Text("required") = "true" Then
                row.SetText "required", Label("yes")
            Else
                row.SetText "required", ""
            End If
            row.SetText "constraints", Constraints(h.Child("schema"))
            row.SetText "description", TrimText(h.Text("description"))
            row.SetText "example", ExampleText(h.Child("example"))
            rows.Push row
        Next i
    End If
    Set HeaderRows = rows
End Function

' ---------------------------------------------------------------------------
' Schemas -> rows
' ---------------------------------------------------------------------------

' Rows that describe a request / response body.
Private Function BodyRows(ByVal schema As clsNode) As clsNode
    Dim rows As clsNode
    Set rows = NewSeq()
    If schema Is Nothing Then
        Set BodyRows = rows
        Exit Function
    End If
    If IsObjectLike(schema) Then
        Flatten rows, schema, "", 0, "|"
    Else
        ' arrays and primitives get a row for the body itself
        AddField rows, "", schema, False, 0, "|"
    End If
    Set BodyRows = rows
End Function

' Emit rows for the children of an object-like schema.
Private Sub Flatten(ByVal rows As clsNode, ByVal schema As clsNode, ByVal prefix As String, _
                    ByVal depth As Long, ByVal visiting As String)
    Dim name As String, s As clsNode, props As clsNode, ap As clsNode, i As Long
    Dim alts As clsNode, kinds As Variant, k As Long, row As clsNode, alt As clsNode, altName As String

    If depth > MAX_DEPTH Then Exit Sub
    name = RefName(schema)
    If name <> "" Then visiting = visiting & name & "|"
    Set s = Merge(Resolve(schema))

    Set props = s.Child("properties")
    If Not props Is Nothing Then
        For i = 1 To props.Count
            AddField rows, JoinName(prefix, props.KeyAt(i)), props.ItemAt(i), IsRequired(s, props.KeyAt(i)), depth, visiting
        Next i
    End If

    Set ap = s.Child("additionalProperties")
    If Not ap Is Nothing Then
        If ap.IsMap Then
            AddField rows, JoinName(prefix, "{key}"), ap, False, depth, visiting
        ElseIf ap.IsScalar And ap.Value = "true" Then
            Set row = NewRow(JoinName(prefix, "{key}"), "any", "", "", "", "")
            rows.Push row
        End If
    End If

    kinds = Array("oneOf", "anyOf")
    For k = 0 To 1
        Set alts = s.Child(kinds(k))
        If Not alts Is Nothing Then
            rows.Push NewRow(DisplayName(prefix), CStr(kinds(k)), "", DiscriminatorText(s), Label(CStr(kinds(k))), "")
            For i = 1 To alts.Count
                Set alt = alts.ItemAt(i)
                altName = RefName(alt)
                If altName = "" Then altName = "#" & i
                rows.Push NewRow(DisplayName(prefix) & " <" & altName & ">", altName, "", _
                                 DiscriminatorValue(s, alt), TrimText(Resolve(alt).Text("description")), "")
                If IsObjectLike(alt) Then Flatten rows, alt, prefix, depth + 1, visiting
            Next i
        End If
    Next k
End Sub

' Emit a row for one field, then rows for its children.
Private Sub AddField(ByVal rows As clsNode, ByVal fieldName As String, ByVal schema As clsNode, _
                     ByVal required As Boolean, ByVal depth As Long, ByVal visiting As String)
    Dim inner As clsNode, s As clsNode, row As clsNode, items As clsNode, target As clsNode
    Dim reqText As String, desc As String, c As String, childPrefix As String

    Set inner = Unwrap(schema)
    Set s = Resolve(inner)

    If required Then reqText = Label("yes")
    desc = TrimText(Resolve(schema).Text("description"))
    If desc = "" Then desc = TrimText(s.Text("description"))
    c = Constraints(schema)

    ' Find the object whose children should follow this row.
    If s.Text("type") = "array" Then
        Set items = s.Child("items")
        If Not items Is Nothing Then
            If IsObjectLike(items) Then
                Set target = items
                childPrefix = fieldName & "[]"
            End If
        End If
    ElseIf IsObjectLike(inner) Then
        Set target = inner
        childPrefix = fieldName
    End If

    If Not target Is Nothing Then
        If RefName(target) <> "" And InStr(1, visiting, "|" & RefName(target) & "|", vbBinaryCompare) > 0 Then
            c = AddLine(c, Label("recursive"))
            Set target = Nothing
        End If
    End If

    Set row = NewRow(DisplayName(fieldName), TypeText(schema), reqText, c, desc, ExampleText(s.Child("example")))
    rows.Push row

    If Not target Is Nothing And depth < MAX_DEPTH Then
        Flatten rows, target, childPrefix, depth + 1, visiting
    End If
End Sub

' "allOf: [ $ref ]" with nothing else meaningful -> the referenced schema.
Private Function Unwrap(ByVal schema As clsNode) As clsNode
    Dim s As clsNode, all As clsNode
    Set Unwrap = schema
    If schema Is Nothing Then Exit Function
    If RefName(schema) <> "" Then Exit Function
    If Not schema.IsMap Then Exit Function
    Set all = schema.Child("allOf")
    If all Is Nothing Then Exit Function
    If all.Count = 1 And schema.Child("properties") Is Nothing Then Set Unwrap = all.ItemAt(1)
End Function

Private Function IsObjectLike(ByVal schema As clsNode) As Boolean
    Dim s As clsNode, ap As clsNode
    If schema Is Nothing Then Exit Function
    Set s = Resolve(Unwrap(schema))
    If Not s.IsMap Then Exit Function
    If s.Text("type") = "array" Then Exit Function
    If s.Text("type") = "object" Then
        IsObjectLike = True
    ElseIf s.Has("properties") Or s.Has("allOf") Or s.Has("oneOf") Or s.Has("anyOf") Then
        IsObjectLike = True
    Else
        Set ap = s.Child("additionalProperties")
        If Not ap Is Nothing Then IsObjectLike = ap.IsMap
    End If
End Function

' Merge allOf parts into one map with properties / required.
Private Function Merge(ByVal s As clsNode) As clsNode
    Dim m As clsNode, props As clsNode, req As clsNode, all As clsNode, part As clsNode
    Dim i As Long, j As Long, src As clsNode

    If Not s.IsMap Then
        Set Merge = s
        Exit Function
    End If
    Set all = s.Child("allOf")
    If all Is Nothing Then
        Set Merge = s
        Exit Function
    End If

    Set m = NewMap()
    Set props = NewMap()
    Set req = NewSeq()
    For i = 1 To all.Count + 1
        If i <= all.Count Then
            Set part = Merge(Resolve(all.ItemAt(i)))
        Else
            Set part = s
        End If
        Set src = part.Child("properties")
        If Not src Is Nothing Then
            For j = 1 To src.Count
                props.SetItem src.KeyAt(j), src.ItemAt(j)
            Next j
        End If
        Set src = part.Child("required")
        If Not src Is Nothing Then
            For j = 1 To src.Count
                req.Push src.ItemAt(j)
            Next j
        End If
        For j = 1 To part.Count
            Select Case part.KeyAt(j)
                Case "allOf", "properties", "required"
                Case Else
                    m.SetItem part.KeyAt(j), part.ItemAt(j)
            End Select
        Next j
    Next i
    m.SetItem "properties", props
    m.SetItem "required", req
    Set Merge = m
End Function

Private Function IsRequired(ByVal s As clsNode, ByVal key As String) As Boolean
    Dim req As clsNode, i As Long
    Set req = s.Child("required")
    If req Is Nothing Then Exit Function
    If Not req.IsSeq Then Exit Function
    For i = 1 To req.Count
        If req.ItemAt(i).Value = key Then
            IsRequired = True
            Exit Function
        End If
    Next i
End Function

Public Function TypeText(ByVal schema As clsNode) As String
    Dim name As String, s As clsNode, t As String, f As String, all As clsNode

    If schema Is Nothing Then Exit Function
    name = RefName(schema)
    Set s = Resolve(schema)
    If Not s.IsMap Then Exit Function
    t = s.Text("type")

    If t = "array" Then
        TypeText = TypeText(s.Child("items")) & "[]"
        Exit Function
    End If
    If t = "" Or t = "object" Then
        Set all = s.Child("allOf")
        If name = "" And Not all Is Nothing Then
            If all.Count = 1 Then
                TypeText = TypeText(all.ItemAt(1))
                Exit Function
            End If
        End If
        If name <> "" Then
            TypeText = name
            Exit Function
        End If
        If s.Has("oneOf") Then
            t = "oneOf"
        ElseIf s.Has("anyOf") Then
            t = "anyOf"
        ElseIf t = "" And Not (s.Has("properties") Or s.Has("allOf") Or s.Has("additionalProperties")) Then
            t = "any"
        Else
            t = "object"
        End If
    End If
    f = s.Text("format")
    If f <> "" Then t = t & "(" & f & ")"
    If name <> "" Then t = name & " : " & t
    TypeText = t
End Function

Public Function Constraints(ByVal schema As clsNode) As String
    Dim c As String, s As clsNode, inner As clsNode, items As clsNode
    If schema Is Nothing Then Exit Function
    c = ConstraintsOf(schema)
    Set inner = Unwrap(schema)
    If Not inner Is schema Then c = JoinLines(c, ConstraintsOf(Resolve(inner)))
    Set s = Resolve(inner)
    If s.Text("type") = "array" Then
        Set items = s.Child("items")
        If Not items Is Nothing Then
            If Not IsObjectLike(items) Then
                If ConstraintsOf(items) <> "" Then c = AddLine(c, "items: " & Replace(ConstraintsOf(items), vbLf, ", "))
            End If
        End If
    End If
    Constraints = c
End Function

Private Function ConstraintsOf(ByVal schema As clsNode) As String
    Dim s As clsNode, c As String, keys As Variant, k As Long, v As clsNode
    Set s = Resolve(schema)
    If Not s.IsMap Then Exit Function
    If Not s.Child("enum") Is Nothing Then c = AddLine(c, "enum: " & JoinScalars(s.Child("enum"), ", "))
    keys = Array("minLength", "maxLength", "minimum", "maximum", "exclusiveMinimum", "exclusiveMaximum", _
                 "multipleOf", "pattern", "minItems", "maxItems", "uniqueItems", "minProperties", _
                 "maxProperties", "default")
    For k = LBound(keys) To UBound(keys)
        Set v = s.Child(CStr(keys(k)))
        If Not v Is Nothing Then
            If v.IsScalar Then c = AddLine(c, keys(k) & ": " & v.Value)
        End If
    Next k
    ' flags may sit next to a $ref / allOf wrapper, so read them from the outer node too
    If schema.IsMap Then
        If schema.Text("nullable") = "true" Or s.Text("nullable") = "true" Then c = AddLine(c, "nullable")
        If schema.Text("readOnly") = "true" Or s.Text("readOnly") = "true" Then c = AddLine(c, "readOnly")
        If schema.Text("writeOnly") = "true" Or s.Text("writeOnly") = "true" Then c = AddLine(c, "writeOnly")
        If schema.Text("deprecated") = "true" Or s.Text("deprecated") = "true" Then c = AddLine(c, "deprecated")
    End If
    ConstraintsOf = c
End Function

Private Function DiscriminatorText(ByVal s As clsNode) As String
    Dim d As clsNode
    Set d = s.Child("discriminator")
    If d Is Nothing Then Exit Function
    DiscriminatorText = "discriminator: " & d.Text("propertyName")
End Function

Private Function DiscriminatorValue(ByVal s As clsNode, ByVal alt As clsNode) As String
    Dim d As clsNode, mapping As clsNode, i As Long, ref As String
    Set d = s.Child("discriminator")
    If d Is Nothing Then Exit Function
    ref = alt.Text("$ref")
    Set mapping = d.Child("mapping")
    If Not mapping Is Nothing Then
        For i = 1 To mapping.Count
            If mapping.ItemAt(i).Value = ref Then
                DiscriminatorValue = d.Text("propertyName") & " = " & mapping.KeyAt(i)
                Exit Function
            End If
        Next i
    End If
    If RefName(alt) <> "" Then DiscriminatorValue = d.Text("propertyName") & " = " & RefName(alt)
End Function

' ---------------------------------------------------------------------------
' Small helpers
' ---------------------------------------------------------------------------

Private Function NewRow(ByVal name As String, ByVal typeStr As String, ByVal required As String, _
                        ByVal constraintText As String, ByVal desc As String, ByVal example As String) As clsNode
    Dim row As clsNode
    Set row = NewMap()
    row.SetText "name", name
    row.SetText "in", ""
    row.SetText "type", typeStr
    row.SetText "required", required
    row.SetText "constraints", constraintText
    row.SetText "description", desc
    row.SetText "example", example
    Set NewRow = row
End Function

Private Function JoinName(ByVal prefix As String, ByVal key As String) As String
    If prefix = "" Then
        JoinName = key
    Else
        JoinName = prefix & "." & key
    End If
End Function

Private Function DisplayName(ByVal fieldName As String) As String
    If fieldName = "" Then
        DisplayName = Label("root")
    Else
        DisplayName = fieldName
    End If
End Function

Private Function ExampleText(ByVal n As clsNode) As String
    If n Is Nothing Then Exit Function
    If n.IsScalar Then
        ExampleText = TrimText(n.Value)
    ElseIf n.IsMap Or n.IsSeq Then
        ExampleText = YamlDump(n)
    End If
End Function

Public Function JoinScalars(ByVal n As clsNode, ByVal sep As String) As String
    Dim i As Long, s As String
    If n Is Nothing Then Exit Function
    If Not n.IsSeq Then
        If n.IsScalar Then JoinScalars = n.Value
        Exit Function
    End If
    For i = 1 To n.Count
        If n.ItemAt(i).IsScalar Then
            If s <> "" Then s = s & sep
            s = s & n.ItemAt(i).Value
        End If
    Next i
    JoinScalars = s
End Function

Private Function AddLine(ByVal text As String, ByVal addition As String) As String
    If text = "" Then
        AddLine = addition
    Else
        AddLine = text & vbLf & addition
    End If
End Function

Private Function JoinLines(ByVal a As String, ByVal b As String) As String
    If a = "" Then
        JoinLines = b
    ElseIf b = "" Then
        JoinLines = a
    Else
        JoinLines = a & vbLf & b
    End If
End Function

Public Function TrimText(ByVal s As String) As String
    Do While Right$(s, 1) = vbLf Or Right$(s, 1) = " "
        s = Left$(s, Len(s) - 1)
    Loop
    TrimText = s
End Function

Private Function CleanId(ByVal s As String) As String
    Dim i As Long, ch As String, r As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[A-Za-z0-9]" Then
            r = r & ch
        Else
            r = r & "_"
        End If
    Next i
    CleanId = r
End Function
