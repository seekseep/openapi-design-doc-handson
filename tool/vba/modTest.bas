Attribute VB_Name = "modTest"
Option Explicit

' Developer tests. Called from template/scripts/excel_dev.py through
' AppleScript (Mac) or by hand from the VBA editor.

' Parse a YAML file and compare the canonical dump with the expected file.
' The expected file is produced by PyYAML (BaseLoader) in make_expected.py.
' Returns "OK", or "NG: ..." / "ERROR: ..." with details.
Public Function TestYamlFile(ByVal yamlPath As String, ByVal expectedPath As String) As String
    Dim actual As String, expected As String, i As Long
    On Error GoTo Failed
    actual = YamlDump(YamlParse(ReadUtf8File(yamlPath)))
    expected = ReadUtf8File(expectedPath)
    If Right$(expected, 1) = vbLf Then expected = Left$(expected, Len(expected) - 1)
    If actual = expected Then
        TestYamlFile = "OK"
        Exit Function
    End If
    For i = 1 To Len(actual)
        If Mid$(actual, i, 1) <> Mid$(expected, i, 1) Then Exit For
    Next i
    TestYamlFile = "NG at " & i & vbLf & "  actual:   " & Mid$(actual, IIf(i > 40, i - 40, 1), 120) & vbLf & "  expected: " & Mid$(expected, IIf(i > 40, i - 40, 1), 120)
    Exit Function
Failed:
    TestYamlFile = "ERROR: " & Err.Description
End Function

' Parse a YAML file that must fail. Returns the error message, or "NO ERROR".
Public Function TestYamlError(ByVal yamlPath As String) As String
    On Error GoTo Failed
    YamlParse ReadUtf8File(yamlPath)
    TestYamlError = "NO ERROR"
    Exit Function
Failed:
    TestYamlError = Err.Description
End Function

' Excel for Mac is sandboxed: ask once for access to a folder (for tests).
' Returns "True" when access was granted.
Public Function GrantAccess(ByVal folderPath As String) As String
#If Mac Then
    GrantAccess = CStr(GrantAccessToMultipleFiles(Array(folderPath)))
#Else
    GrantAccess = "True"
#End If
End Function

' Run every fixture next to this workbook (../fixtures) and write the
' results to a new sheet. For checking the parser on Windows by hand:
' open tool/test/parser-test.xlsm and run RunAllTests.
Public Sub RunAllTests()
    Dim fixtureDir As String, f As String, ws As Worksheet, r As Long, result As String
    Dim failed As Long, total As Long

    fixtureDir = JoinPath(JoinPath(ThisWorkbook.Path, ".."), "fixtures")
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Range("A1").Value = "file"
    ws.Range("B1").Value = "result"
    r = 1

    f = Dir(JoinPath(fixtureDir, "*.yaml"))
    Do While f <> ""
        result = TestYamlFile(JoinPath(fixtureDir, f), JoinPath(fixtureDir, Left$(f, Len(f) - 5) & ".expected"))
        r = r + 1
        total = total + 1
        ws.Cells(r, 1).Value = f
        ws.Cells(r, 2).Value = result
        If result <> "OK" Then failed = failed + 1
        f = Dir()
    Loop

    f = Dir(JoinPath(JoinPath(fixtureDir, "errors"), "*.yaml"))
    Do While f <> ""
        result = TestYamlError(JoinPath(JoinPath(fixtureDir, "errors"), f))
        r = r + 1
        total = total + 1
        ws.Cells(r, 1).Value = "errors/" & f
        ws.Cells(r, 2).Value = result
        If result = "NO ERROR" Then failed = failed + 1
        f = Dir()
    Loop

    ws.Columns("A:B").AutoFit
    MsgBox total & " tests, " & failed & " failed"
End Sub
