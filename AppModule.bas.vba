Attribute VB_Name = "AppModule"
Option Explicit

Sub Start�Z���t���r���r���[�A�[()
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

Sub test�r�����()
    Dim bd As Border
    Set bd = Selection.Borders(xlInsideHorizontal)
    Debug.Print bd.LineStyle
End Sub

Rem �ʂ̉�����F�S�Ă̌r���𒆊Ԑ��Ƃ��čēK�p����

