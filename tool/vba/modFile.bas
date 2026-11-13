Attribute VB_Name = "modFile"
Option Explicit

' File helpers that work on both Windows and Mac.
' (No Scripting.FileSystemObject / ADODB.Stream: they do not exist on Mac.)

Public Function PickYamlFile() As String
    Dim r As Variant
#If Mac Then
    r = Application.GetOpenFilename()
#Else
    r = Application.GetOpenFilename("OpenAPI (*.yaml;*.yml),*.yaml;*.yml,All files (*.*),*.*")
#End If
    If VarType(r) = vbBoolean Then
        PickYamlFile = ""
    Else
        PickYamlFile = CStr(r)
    End If
End Function

Public Function FileExists(ByVal path As String) As Boolean
    On Error Resume Next
    FileExists = (Len(Dir(path)) > 0)
    If Err.Number <> 0 Then FileExists = False
    On Error GoTo 0
End Function

Public Function FolderExists(ByVal path As String) As Boolean
    On Error Resume Next
    FolderExists = (Len(Dir(path, vbDirectory)) > 0)
    If Err.Number <> 0 Then FolderExists = False
    On Error GoTo 0
End Function

Public Function JoinPath(ByVal a As String, ByVal b As String) As String
    Dim sep As String
    sep = Application.PathSeparator
    If Right$(a, 1) = sep Then
        JoinPath = a & b
    Else
        JoinPath = a & sep & b
    End If
End Function

Public Function FileName(ByVal path As String) As String
    Dim p As Long
    p = InStrRev(path, Application.PathSeparator)
    FileName = Mid$(path, p + 1)
End Function

' Read a whole file as UTF-8 text (a BOM is skipped).
Public Function ReadUtf8File(ByVal path As String) As String
    Dim f As Integer, n As Long, b() As Byte

#If Mac Then
    ' Excel for Mac is sandboxed. Ask for access (shows a dialog only once).
    On Error Resume Next
    GrantAccessToMultipleFiles Array(path)
    On Error GoTo 0
#End If

    f = FreeFile
    Open path For Binary Access Read As #f
    n = LOF(f)
    If n = 0 Then
        Close #f
        ReadUtf8File = ""
        Exit Function
    End If
    ReDim b(0 To n - 1)
    Get #f, , b
    Close #f
    ReadUtf8File = Utf8Decode(b, n)
End Function

Public Function Utf8Decode(b() As Byte, ByVal n As Long) As String
    Dim buf As String, outPos As Long, i As Long
    Dim c As Long, cp As Long, need As Long, k As Long

    buf = String$(n, " ")
    i = 0
    If n >= 3 Then
        If b(0) = &HEF And b(1) = &HBB And b(2) = &HBF Then i = 3
    End If
    Do While i < n
        c = b(i)
        If c < &H80 Then
            cp = c
            need = 0
        ElseIf (c And &HE0) = &HC0 Then
            cp = c And &H1F
            need = 1
        ElseIf (c And &HF0) = &HE0 Then
            cp = c And &HF
            need = 2
        ElseIf (c And &HF8) = &HF0 Then
            cp = c And &H7
            need = 3
        Else
            cp = 65533
            need = 0
        End If
        If i + need >= n Then
            cp = 65533
            need = n - i - 1
        Else
            For k = 1 To need
                cp = cp * &H40 + (b(i + k) And &H3F)
            Next k
        End If
        i = i + need + 1

        If cp >= 65536 Then
            cp = cp - 65536
            outPos = outPos + 1
            Mid$(buf, outPos, 1) = ChrW(55296 + Int(cp / 1024))
            outPos = outPos + 1
            Mid$(buf, outPos, 1) = ChrW(56320 + (cp And 1023))
        Else
            outPos = outPos + 1
            Mid$(buf, outPos, 1) = ChrW(cp)
        End If
    Loop
    Utf8Decode = Left$(buf, outPos)
End Function
