Attribute VB_Name = "AppMain"
Rem
Rem @appname ExcelCellBorderViewer - セル付き罫線ビュアー
Rem
Rem @module AppMain
Rem
Rem @author @KotorinChunChun
Rem
Rem @update
Rem    2024/03/27 : 暫定版
Rem    2024/03/28 : 副作用解消版(XML対応版)
Rem
Option Explicit
Option Private Module

Public Const APP_NAME = "セル付き罫線ビュアーアドイン"
Public Const APP_CREATER = "@KotorinChunChun"
Public Const APP_VERSION = "0.10"
Public Const APP_UPDATE = "2024/03/28"
Public Const APP_URL = "https://github.com/KotorinChunChun/ExcelCellBorderViewer"

Rem --------------------------------------------------
Rem Global #Const EnableCfgExcelCellBorderViewer=1
Public cfgCellRowCount As Long
Public cfgCellColCount As Long
Public cfgCellGap As Long
Public cfgBorderSize As Long
Rem --------------------------------------------------
Rem アドイン実行時
Sub AddinStart()
    Call Startセル付き罫線ビューアー
End Sub

Rem アドイン一時停止時
Sub AddinStop()
    End
End Sub

Rem アドイン設定表示
Sub AddinConfig(): Call SettingForm.Show: End Sub

Rem アドイン情報表示
Sub AddinInfo()
    Select Case MsgBox(ThisWorkbook.Name & vbLf & vbLf & _
            "バージョン : " & APP_VERSION & vbLf & _
            "更新日　　 : " & APP_UPDATE & vbLf & _
            "開発者　　 : " & APP_CREATER & vbLf & _
            "実行パス　 : " & ThisWorkbook.Path & vbLf & _
            "公開ページ : " & APP_URL & vbLf & _
            vbLf & _
            "使い方や最新版を探しに公開ページを開きますか？" & _
            "", vbInformation + vbYesNo, "バージョン情報")
        Case vbNo
            Rem
        Case vbYes
            CreateObject("Wscript.Shell").Run APP_URL, 3
    End Select
End Sub

Rem アドイン完全終了
Sub AddinEnd(): ThisWorkbook.Close False: End Sub

Rem 設定値の書き込み
Sub WriteSetting(Key As String, Value)
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets(1).ListObjects(1)
    Dim rr As Long
    For rr = 1 To lo.ListRows.Count
        If lo.Range.Cells(1 + rr, 1) = Key Then
            lo.Range.Cells(1 + rr, 2) = Value
            Exit Sub
        End If
    Next
    Rem key not found
    Err.Raise 9999, , "設定テーブルで未定義の設定値を書き込んでいます"
End Sub

Rem 設定値の読み込み
Sub LoadSettings()
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets(1).ListObjects(1)
    Dim dic As New Dictionary
    Dim rr As Long
    For rr = 1 To lo.ListRows.Count
        dic.Add lo.Range.Cells(rr, 1).Value, lo.Range.Cells(rr, 2).Value
    Next
    cfgCellRowCount = dic("cfgCellRowCount")
    cfgCellColCount = dic("cfgCellColCount")
    cfgCellGap = dic("cfgCellGap")
    cfgBorderSize = dic("cfgBorderSize")
End Sub
Rem --------------------------------------------------

Sub Startセル付き罫線ビューアー()
    Call LoadSettings
    
    Rem イベントフックを保持するための保持
    Static fm As CellBorderViewForm
    
    Rem フォームの破棄チェック
    On Error Resume Next
        Debug.Print Now, "fm.Visible : " & fm.Visible
        If Err Then Set fm = Nothing
    On Error GoTo 0
    
    Rem フォームの多重起動の防止と表示
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
