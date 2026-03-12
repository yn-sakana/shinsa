VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShinsaConfig"
Option Explicit

Public SelfAddress As String
Public MailFolder As String      ' メールが保存されたフォルダのパス
Public CaseFolder As String      ' 案件フォルダのルートパス
Public KeyColumn As String       ' 案件テーブルのキー列名
Public MailLinkColumn As String  ' メール紐づけ用の列名（メールアドレス等）
