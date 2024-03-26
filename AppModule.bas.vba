Attribute VB_Name = "AppModule"
Option Explicit

Sub Startセル付き罫線ビューアー()
    Static fm As CellBorderViewForm
    
    On Error Resume Next
        Debug.Print Now, "fm.Visible : " & fm.Visible
        If Err Then Set fm = Nothing
    On Error GoTo 0
    
    If fm Is Nothing Then
        Set fm = New CellBorderViewForm
        fm.Show False
    End If
End Sub

Sub test罫線情報()
    Dim bd As Border
    Set bd = Selection.Borders(xlInsideHorizontal)
    Debug.Print bd.LineStyle
End Sub

Rem 別の解決策：全ての罫線を中間線として再適用する

