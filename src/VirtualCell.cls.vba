VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VirtualCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Rem VirtualCell : �Z���̏�Ԃ��t�H�[����̃R���g���[���ɉ��z�I�ɍČ����邽�߂̃N���X

Rem 5��ނ̃Z����\������R���g���[��
Public WithEvents LabelCell As MSForms.Label
Attribute LabelCell.VB_VarHelpID = -1
Public WithEvents ImageLeft As MSForms.Image
Attribute ImageLeft.VB_VarHelpID = -1
Public WithEvents ImageTop As MSForms.Image
Attribute ImageTop.VB_VarHelpID = -1
Public WithEvents ImageRight As MSForms.Image
Attribute ImageRight.VB_VarHelpID = -1
Public WithEvents ImageBottom As MSForms.Image
Attribute ImageBottom.VB_VarHelpID = -1

Rem �r���̐F�̒�`
Function RgbNone(): RgbNone = RGB(230, 230, 230): End Function
Function RgbContinuous(): RgbContinuous = RGB(0, 0, 0): End Function
    
Rem �w�肵���Z���̏���ێ����Ă���R���g���[���ɔ��f����
Public Sub Update(rng As Range)
    If rng.Rows.Count > 1 Or rng.Columns.Count > 1 Then MsgBox "�Z���͈͂ɂ͔�Ή�": End
    
    Dim BorderInfo As Object: Set BorderInfo = GetBorderInfo(rng)
    
    SwitchImageStyle Me.ImageLeft, BorderInfo("Left")
    SwitchImageStyle Me.ImageTop, BorderInfo("Top")
    SwitchImageStyle Me.ImageRight, BorderInfo("Right")
    SwitchImageStyle Me.ImageBottom, BorderInfo("Bottom")
    LabelCell.Caption = rng.Address(False, False)
    LabelCell.Font.Size = IIf(Len(LabelCell.Caption) <= 4, 18, 9)
End Sub

Rem �Z���E�㉺���E�̌r���̑S�Ă𖳌�������i���[�N�V�[�g�O�̃Z����\������ꍇ�Ɏg�p�j
Public Sub GrayOut()
    Me.ImageLeft.BackColor = RgbNone
    Me.ImageTop.BackColor = RgbNone
    Me.ImageRight.BackColor = RgbNone
    Me.ImageBottom.BackColor = RgbNone
    LabelCell.Caption = ""
End Sub

Rem �w�肵���C���[�W�R���g���[���̌����ڂ�Border�̗L���ɍ��킹�ĕω�������
Private Sub SwitchImageStyle(img As MSForms.Image, existsBorder As Boolean)
    img.BackColor = IIf(existsBorder, RgbContinuous, RgbNone)
    img.BorderStyle = fmBorderStyleNone
End Sub

Rem �w�肵���Z���ɕt�^���ꂽ�㉺���E�̐^�̌r������XML����ǂݎ��
Private Function GetBorderInfo(rng As Range) As Object
    Dim BorderInfo As Object: Set BorderInfo = CreateObject("Scripting.Dictionary")
    Dim BorderName: For Each BorderName In VBA.Array("Top", "Bottom", "Left", "Right"): BorderInfo(BorderName) = False: Next
    With CreateObject("MSXML2.DOMDocument")
        Call .LoadXML(rng.Value(xlRangeValueXMLSpreadsheet))
        Dim Node As Object
        For Each Node In .SelectNodes("//Style[not(@ss:ID='Default')]/Borders/Border")
            BorderInfo(Node.Attributes.getNamedItem("ss:Position").Text) = True
        Next
    End With
    Set GetBorderInfo = BorderInfo
End Function

Private Sub LabelCell_Click()
    If LabelCell.Caption = "" Then Exit Sub
    Dim rng As Range
    Set rng = ActiveSheet.Range(LabelCell.Caption)
    rng.Activate
    AppActivate ActiveWindow.Application.Caption
End Sub

Private Sub ImageTop_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If LabelCell.Caption = "" Then Exit Sub
    Dim rng As Range
    Set rng = ActiveSheet.Range(LabelCell.Caption)
    Call SwitchCellBorderLine(rng.Borders(xlTop), ImageTop.BackColor = RgbNone)
    Call Update(rng)
    AppActivate ActiveWindow.Application.Caption
End Sub

Private Sub ImageLeft_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If LabelCell.Caption = "" Then Exit Sub
    Dim rng As Range
    Set rng = ActiveSheet.Range(LabelCell.Caption)
    Call SwitchCellBorderLine(rng.Borders(xlLeft), ImageLeft.BackColor = RgbNone)
    Call Update(rng)
    AppActivate ActiveWindow.Application.Caption
End Sub

Private Sub ImageBottom_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If LabelCell.Caption = "" Then Exit Sub
    Dim rng As Range
    Set rng = ActiveSheet.Range(LabelCell.Caption)
    Call SwitchCellBorderLine(rng.Borders(xlBottom), ImageBottom.BackColor = RgbNone)
    Call Update(rng)
    AppActivate ActiveWindow.Application.Caption
End Sub

Private Sub ImageRight_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If LabelCell.Caption = "" Then Exit Sub
    Dim rng As Range
    Set rng = ActiveSheet.Range(LabelCell.Caption)
    Call SwitchCellBorderLine(rng.Borders(xlRight), ImageRight.BackColor = RgbNone)
    Call Update(rng)
    AppActivate ActiveWindow.Application.Caption
End Sub

Private Function SwitchCellBorderLine(bd As Excel.Border, tf As Boolean) As Boolean
    If tf Then
        bd.LineStyle = xlContinuous
        bd.ColorIndex = 0
        bd.TintAndShade = 0
        bd.Weight = xlThin
    Else
        bd.LineStyle = xlNone
    End If
    SwitchCellBorderLine = tf
End Function
