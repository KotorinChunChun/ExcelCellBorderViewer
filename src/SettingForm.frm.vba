VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SettingForm 
   Caption         =   "設定"
   ClientHeight    =   6180
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4890
   OleObjectBlob   =   "SettingForm.frm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "SettingForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
    
#If EnableCfgExcelCellBorderViewer = 1 Then
    Rem 設定値の読み込み（雑な実装）
'    On Error Resume Next
    WriteSetting "cfgCellRowCount", SpinButtonRowCount.Value
    WriteSetting "cfgCellColCount", SpinButtonColCount.Value
    WriteSetting "cfgCellGap", SpinButtonCellGap.Value
    WriteSetting "cfgBorderSize", SpinButtonBorderSize.Value
    On Error GoTo 0
#Else
    MsgBox "条件付きコンパイルにEnableCfgExcelCellBorderViewer宣言が足りません"
    End
#End If
    
    Unload Me
End Sub

Private Sub SpinButtonBorderSize_Change()
    TextBoxBorderSize.Text = SpinButtonBorderSize.Value
End Sub

Private Sub SpinButtonCellGap_Change()
    TextBoxCellGap.Text = SpinButtonCellGap.Value
End Sub

Private Sub SpinButtonColCount_Change()
    TextBoxColCount.Text = SpinButtonColCount.Value
End Sub

Private Sub SpinButtonRowCount_Change()
    TextBoxRowCount.Text = SpinButtonRowCount.Value
End Sub

Private Sub UserForm_Initialize()
    Me.Caption = APP_NAME & " - 設定"
    SpinButtonRowCount.Value = 3
    SpinButtonColCount.Value = 3
    SpinButtonBorderSize.Value = 3
    SpinButtonCellGap.Value = 3
End Sub
