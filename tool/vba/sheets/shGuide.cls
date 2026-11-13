Option Explicit

' The "button" on the guide sheet is a cell with a hyperlink to itself.
' (A cell is used instead of a form button so that the workbook can be
' built without Excel.)
Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
    If Target.Range.Address = Me.Range("ImportButton").Address Then ImportOpenApi
End Sub
