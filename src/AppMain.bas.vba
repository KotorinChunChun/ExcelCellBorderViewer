Attribute VB_Name = "AppMain"
Rem
Rem @appname ExcelCellBorderViewer - �Z���t���r���r���A�[
Rem
Rem @module AppMain
Rem
Rem @author @KotorinChunChun
Rem
Rem @update
Rem    2024/03/27 : �b���
Rem    2024/03/28 : ����p������(XML�Ή���)
Rem
Option Explicit
Option Private Module

Public Const APP_NAME = "�Z���t���r���r���A�[�A�h�C��"
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
Rem �A�h�C�����s��
Sub AddinStart()
    Call Start�Z���t���r���r���[�A�[
End Sub

Rem �A�h�C���ꎞ��~��
Sub AddinStop()
    End
End Sub

Rem �A�h�C���ݒ�\��
Sub AddinConfig(): Call SettingForm.Show: End Sub

Rem �A�h�C�����\��
Sub AddinInfo()
    Select Case MsgBox(ThisWorkbook.Name & vbLf & vbLf & _
            "�o�[�W���� : " & APP_VERSION & vbLf & _
            "�X�V���@�@ : " & APP_UPDATE & vbLf & _
            "�J���ҁ@�@ : " & APP_CREATER & vbLf & _
            "���s�p�X�@ : " & ThisWorkbook.Path & vbLf & _
            "���J�y�[�W : " & APP_URL & vbLf & _
            vbLf & _
            "�g������ŐV�ł�T���Ɍ��J�y�[�W���J���܂����H" & _
            "", vbInformation + vbYesNo, "�o�[�W�������")
        Case vbNo
            Rem
        Case vbYes
            CreateObject("Wscript.Shell").Run APP_URL, 3
    End Select
End Sub

Rem �A�h�C�����S�I��
Sub AddinEnd(): ThisWorkbook.Close False: End Sub

Rem �ݒ�l�̏�������
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
    Err.Raise 9999, , "�ݒ�e�[�u���Ŗ���`�̐ݒ�l����������ł��܂�"
End Sub

Rem �ݒ�l�̓ǂݍ���
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

Sub Start�Z���t���r���r���[�A�[()
    Call LoadSettings
    
    Rem �C�x���g�t�b�N��ێ����邽�߂̕ێ�
    Static fm As CellBorderViewForm
    
    Rem �t�H�[���̔j���`�F�b�N
    On Error Resume Next
        Debug.Print Now, "fm.Visible : " & fm.Visible
        If Err Then Set fm = Nothing
    On Error GoTo 0
    
    Rem �t�H�[���̑��d�N���̖h�~�ƕ\��
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
