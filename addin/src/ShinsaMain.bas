Attribute VB_Name = "ShinsaMain"
Option Explicit

' Shinsa Excel Add-in
' 案件テーブルの選択行に連動して、関連メール・フォルダを表示する

Private Const CONFIG_SHEET_NAME As String = "_shinsa_config"

' --- Entry Points (Ribbon / Shortcut) ---

Public Sub Shinsa_ShowPanel()
    frmCaseDetail.Show vbModeless
End Sub

Public Sub Shinsa_ImportMail()
    Dim cfg As ShinsaConfig
    Set cfg = LoadConfig()
    If Len(cfg.MailFolder) = 0 Then
        MsgBox "メールフォルダが設定されていません。設定画面で指定してください。", vbExclamation
        Exit Sub
    End If
    If Len(cfg.SelfAddress) = 0 Then
        MsgBox "自分のメールアドレスが設定されていません。", vbExclamation
        Exit Sub
    End If

    Dim count As Long
    count = ShinsaMailImport.ImportMail(cfg)
    MsgBox count & " 件のメールをインポートしました。", vbInformation
End Sub

Public Sub Shinsa_ShowSettings()
    frmSettings.Show vbModal
End Sub

' --- Config ---

Public Function LoadConfig() As ShinsaConfig
    Dim cfg As New ShinsaConfig
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        ' Defaults
        cfg.SelfAddress = ""
        cfg.MailFolder = ""
        cfg.CaseFolder = ""
        cfg.KeyColumn = ""
        cfg.MailLinkColumn = ""
    Else
        cfg.SelfAddress = CStr(ws.Range("B1").Value)
        cfg.MailFolder = CStr(ws.Range("B2").Value)
        cfg.CaseFolder = CStr(ws.Range("B3").Value)
        cfg.KeyColumn = CStr(ws.Range("B4").Value)
        cfg.MailLinkColumn = CStr(ws.Range("B5").Value)
    End If

    Set LoadConfig = cfg
End Function

Public Sub SaveConfig(cfg As ShinsaConfig)
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = CONFIG_SHEET_NAME
        ws.Visible = xlSheetVeryHidden
        ws.Range("A1").Value = "self_address"
        ws.Range("A2").Value = "mail_folder"
        ws.Range("A3").Value = "case_folder"
        ws.Range("A4").Value = "key_column"
        ws.Range("A5").Value = "mail_link_column"
    End If

    ws.Range("B1").Value = cfg.SelfAddress
    ws.Range("B2").Value = cfg.MailFolder
    ws.Range("B3").Value = cfg.CaseFolder
    ws.Range("B4").Value = cfg.KeyColumn
    ws.Range("B5").Value = cfg.MailLinkColumn
End Sub

' --- Selection Helper ---

Public Function GetSelectedKeyValue() As String
    Dim cell As Range
    Set cell = Selection
    If cell Is Nothing Then Exit Function

    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = cell.ListObject
    On Error GoTo 0
    If tbl Is Nothing Then Exit Function

    Dim cfg As ShinsaConfig
    Set cfg = LoadConfig()
    If Len(cfg.KeyColumn) = 0 Then Exit Function

    ' Find key column in table
    Dim col As ListColumn
    On Error Resume Next
    Set col = tbl.ListColumns(cfg.KeyColumn)
    On Error GoTo 0
    If col Is Nothing Then Exit Function

    ' Get key value from same row
    Dim rowOffset As Long
    rowOffset = cell.Row - tbl.DataBodyRange.Row + 1
    If rowOffset < 1 Or rowOffset > tbl.DataBodyRange.Rows.Count Then Exit Function

    GetSelectedKeyValue = CStr(tbl.DataBodyRange.Cells(rowOffset, col.Index).Value)
End Function
