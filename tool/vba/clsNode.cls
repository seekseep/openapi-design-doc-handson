VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsNode"
Option Explicit

' A node of a parsed YAML document (also used for the view model).
' Kind: 0 = null, 1 = scalar, 2 = map, 3 = sequence.
' Maps keep insertion order. Lookups are linear, which is fine for
' the size of an OpenAPI document.

Public Kind As Long
Public Value As String
Public SrcLine As Long

Private mKeys() As String
Private mItems() As clsNode
Private mCount As Long
Private mCapacity As Long

Public Property Get Count() As Long
    Count = mCount
End Property

Public Function IsScalar() As Boolean
    IsScalar = (Kind = 1)
End Function

Public Function IsMap() As Boolean
    IsMap = (Kind = 2)
End Function

Public Function IsSeq() As Boolean
    IsSeq = (Kind = 3)
End Function

Private Sub Grow()
    If mCapacity = 0 Then
        mCapacity = 8
        ReDim mKeys(1 To mCapacity)
        ReDim mItems(1 To mCapacity)
    ElseIf mCount >= mCapacity Then
        mCapacity = mCapacity * 2
        ReDim Preserve mKeys(1 To mCapacity)
        ReDim Preserve mItems(1 To mCapacity)
    End If
End Sub

Public Function IndexOf(ByVal key As String) As Long
    Dim i As Long
    For i = 1 To mCount
        If mKeys(i) = key Then
            IndexOf = i
            Exit Function
        End If
    Next i
    IndexOf = 0
End Function

Public Function Has(ByVal key As String) As Boolean
    Has = (IndexOf(key) > 0)
End Function

' Map: add or replace.
Public Sub SetItem(ByVal key As String, ByVal node As clsNode)
    Dim i As Long
    i = IndexOf(key)
    If i > 0 Then
        Set mItems(i) = node
        Exit Sub
    End If
    Grow
    mCount = mCount + 1
    mKeys(mCount) = key
    Set mItems(mCount) = node
End Sub

' Map: set a scalar value.
Public Sub SetText(ByVal key As String, ByVal text As String)
    Dim n As New clsNode
    n.Kind = 1
    n.Value = text
    SetItem key, n
End Sub

' Sequence: append.
Public Sub Push(ByVal node As clsNode)
    Grow
    mCount = mCount + 1
    mKeys(mCount) = ""
    Set mItems(mCount) = node
End Sub

Public Function KeyAt(ByVal i As Long) As String
    KeyAt = mKeys(i)
End Function

Public Function ItemAt(ByVal i As Long) As clsNode
    Set ItemAt = mItems(i)
End Function

' Map: child node or Nothing.
Public Function Child(ByVal key As String) As clsNode
    Dim i As Long
    If Kind <> 2 Then
        Set Child = Nothing
        Exit Function
    End If
    i = IndexOf(key)
    If i = 0 Then
        Set Child = Nothing
    Else
        Set Child = mItems(i)
    End If
End Function

' Map: scalar text of a child, or the default.
Public Function Text(ByVal key As String, Optional ByVal def As String = "") As String
    Dim n As clsNode
    Set n = Child(key)
    If n Is Nothing Then
        Text = def
    ElseIf n.Kind = 1 Then
        Text = n.Value
    Else
        Text = def
    End If
End Function
