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
Rem 方針：VBAでも上のセルの下線か、下のセルの上線かを識別はできない
Rem セルを別のところにコピーすれば正しい判定ができる

Public Cell As MSForms.Label
Public Left As MSForms.Image
Public Top As MSForms.Image
Public Right As MSForms.Image
Public Bottom As MSForms.Image

Function RgbNone(): RgbNone = RGB(230, 230, 230): End Function

Rem 指定したセルの情報を保持しているコントロールに反映する
Public Sub Update(rng As Range)
    If rng.Rows.Count > 1 Or rng.Columns.Count > 1 Then MsgBox "セル範囲には非対応"
    
    Rem ダミーのセルにコピーしてから判定することで確認できる
    Dim dummyCell As Range
    Set dummyCell = Range("AAA100")
    rng.Copy dummyCell
    
    SwitchImageStyle Me.Left, dummyCell.Borders(xlEdgeLeft)
    SwitchImageStyle Me.Top, dummyCell.Borders(xlEdgeTop)
    SwitchImageStyle Me.Right, dummyCell.Borders(xlEdgeRight)
    SwitchImageStyle Me.Bottom, dummyCell.Borders(xlEdgeBottom)
    Cell.Caption = rng.Address(False, False)
End Sub

Public Sub GrayOut()
    Me.Left.BackColor = RgbNone
    Me.Top.BackColor = RgbNone
    Me.Right.BackColor = RgbNone
    Me.Bottom.BackColor = RgbNone
    Cell.Caption = ""
End Sub

Rem
Public Sub SwitchImageStyle(img As MSForms.Image, bd As Excel.Border)
    img.BackColor = IIf(bd.LineStyle = xlNone, RgbNone, RGB(0, 0, 0))
    img.BorderStyle = fmBorderStyleNone
End Sub

