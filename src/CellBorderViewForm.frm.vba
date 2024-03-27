VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CellBorderViewForm 
   Caption         =   "UserForm1"
   ClientHeight    =   5955
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10455
   OleObjectBlob   =   "CellBorderViewForm.frm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "CellBorderViewForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Rem
Rem  セル付き罫線ビュアー - ExcelCellBorderViewer
Rem
Rem    作者  ことりちゅん - KotorinChunChun
Rem     URL  https://twitter.com/KotorinChunChun
Rem
Rem    更新  2024/3/27
Rem
Rem  公開先  https://github.com/KotorinChunChun/ExcelCellBorderViewer
Rem
Rem
Rem このクラスを活用するには、以下のコードを標準モジュールに貼って起動ボタンに設定してください
'Sub Startセル付き罫線ビューアー()
'    Static fm As CellBorderViewForm
'    Set fm = New CellBorderViewForm
'    fm.Show False
'End Sub

Private VCells() As VirtualCell
Private WithEvents app As Excel.Application
Attribute app.VB_VarHelpID = -1

Const defCellRowCount = 3
Const defCellColCount = 3
Const CellHeight = 6 * 3
Const CellWidth = 6 * 8
Const defCellGap = 3
Const defBorderSize = 3
Const rgbCenterBackColor = 200 + 200 * 2 ^ 8 + 255 * 2 ^ 16

Public CellRowCount As Long
Public CellColCount As Long
Public CellGap As Long
Public BorderSize As Long

Private Sub app_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    Dim rr As Long, cc As Long
    For rr = 1 To CellRowCount
        For cc = 1 To CellColCount
            Dim offsetR As Long, offsetC As Long
            offsetR = rr - Int(CellRowCount / 2)
            offsetC = cc - Int(CellColCount / 2)
            
            On Error Resume Next
            Dim TargetCell As Range
            Set TargetCell = Nothing
            Set TargetCell = Target.Cells(offsetR, offsetC)
            On Error GoTo 0
            If TargetCell Is Nothing Then
                Call VCells(rr, cc).GrayOut
            Else
                Rem 指定したセルの情報を保持しているコントロールに反映する
                Call VCells(rr, cc).Update(TargetCell)
            End If
        Next
    Next
    
End Sub

Private Sub UserForm_Initialize()
    Set app = Application
    
#If EnableCfgExcelCellBorderViewer = 1 Then
    CellRowCount = cfgCellRowCount
    CellColCount = cfgCellColCount
    CellGap = cfgCellGap
    BorderSize = cfgBorderSize
#End If
    If CellRowCount = 0 Then CellRowCount = defCellRowCount
    If CellColCount = 0 Then CellColCount = defCellColCount
    If CellGap = 0 Then CellGap = defCellGap
    If BorderSize = 0 Then BorderSize = defBorderSize
    
    Me.Caption = "セル付き罫線ビューアー"
    Const CtrlToFormPixel = 2.54 '≒567:1440 コントロールとフォームのピクセルの比率
    Me.Height = (CellHeight + CellGap + 1 + 4) * CellRowCount * CtrlToFormPixel
    Me.Width = (CellWidth + CellGap + 1 - 4) * CellColCount * CtrlToFormPixel
    
    ReDim VCells(1 To CellRowCount, 1 To CellColCount)
    
    Dim rr As Long, cc As Long
    For rr = 1 To CellRowCount
        For cc = 1 To CellColCount
            Set VCells(rr, cc) = New VirtualCell
            With VCells(rr, cc)
                Dim ctr As MSForms.control
                
                Set .LabelCell = Me.Controls.Add("Forms.Label.1")
                Set ctr = .LabelCell
                With ctr
                    .Height = CellHeight
                    .Width = CellWidth
                    .Top = 5 + (rr - 1) * (CellHeight + CellGap * 3) + CellGap * 1
                    .Left = 5 + (cc - 1) * (CellWidth + CellGap * 3) + CellGap * 1
                End With
                .LabelCell.Font.Size = CellHeight
                .LabelCell.Caption = "AA123"  '文字幅確認のための適当なセルアドレス
                .LabelCell.TextAlign = fmTextAlignCenter
                
                Set .ImageLeft = Me.Controls.Add("Forms.Image.1")
                Set ctr = .ImageLeft
                With ctr
                    .Height = CellHeight
                    .Width = BorderSize
                    .Top = 5 + (rr - 1) * (CellHeight + CellGap * 3) + CellGap * 1
                    .Left = 5 + (cc - 1) * (CellWidth + CellGap * 3)
                End With
                
                Set .ImageRight = Me.Controls.Add("Forms.Image.1")
                Set ctr = .ImageRight
                With ctr
                    .Height = CellHeight
                    .Width = BorderSize
                    .Top = 5 + (rr - 1) * (CellHeight + CellGap * 3) + CellGap * 1
                    .Left = 5 + (cc - 0) * (CellWidth + CellGap * 3) - CellGap * 2
                End With
                
                Set .ImageTop = Me.Controls.Add("Forms.Image.1")
                Set ctr = .ImageTop
                With ctr
                    .Height = BorderSize
                    .Width = CellWidth
                    .Top = 5 + (rr - 1) * (CellHeight + CellGap * 3)
                    .Left = 5 + (cc - 1) * (CellWidth + CellGap * 3) + CellGap * 1
                End With
                
                Set .ImageBottom = Me.Controls.Add("Forms.Image.1")
                Set ctr = .ImageBottom
                With ctr
                    .Height = BorderSize
                    .Width = CellWidth
                    .Top = 5 + (rr - 0) * (CellHeight + CellGap * 3) - CellGap * 2
                    .Left = 5 + (cc - 1) * (CellWidth + CellGap * 3) + CellGap * 1
                End With
                
            End With
        Next
    Next
    
    VCells(1 + Int(CellRowCount / 2), 1 + Int(CellColCount / 2)).LabelCell.BackColor = rgbCenterBackColor
    
    Rem 現状を描画
    Call app_SheetSelectionChange(ActiveSheet, ActiveCell)
    
End Sub
