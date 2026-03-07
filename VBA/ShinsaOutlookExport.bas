Attribute VB_Name = "ShinsaOutlookExport"
Option Explicit

Private Const OLMSGUNICODE As Long = 9

Public Function Shinsa_ExportMail(ByVal exportRoot As String, ByVal stateFilePath As String, ByVal selfAddress As String) As Long
    On Error GoTo ErrHandler

    Dim exported As Object
    Dim store As Outlook.Store
    Dim accountSmtp As String
    Dim exportedCount As Long

    If Len(exportRoot) = 0 Then
        Err.Raise vbObjectError + 513, , "exportRoot is empty"
    End If
    If Len(selfAddress) = 0 Then
        Err.Raise vbObjectError + 513, , "selfAddress is empty"
    End If

    EnsureFolder exportRoot
    Set exported = LoadExportedIds(stateFilePath)

    exportedCount = 0
    For Each store In Application.Session.Stores
        accountSmtp = GetStoreSmtpAddress(store)
        If Len(accountSmtp) > 0 Then
            If LCase$(accountSmtp) = LCase$(selfAddress) Then
                exportedCount = exportedCount + ExportFolderTree(store.GetRootFolder, exportRoot, accountSmtp, exported)
            End If
        End If
    Next store

    SaveExportedIds stateFilePath, exported

    Shinsa_ExportMail = exportedCount
    Exit Function

ErrHandler:
    MsgBox "Shinsa Outlook export failed: " & Err.Description, vbExclamation
    Shinsa_ExportMail = 0
End Function

Private Function ExportFolderTree(ByVal targetFolder As Outlook.Folder, ByVal exportRoot As String, ByVal accountSmtp As String, ByVal exported As Object) As Long
    On Error GoTo FolderError

    Dim folderRoot As String
    Dim items As Outlook.Items
    Dim itemIndex As Long
    Dim currentItem As Object
    Dim mail As Outlook.MailItem
    Dim child As Outlook.Folder
    Dim total As Long

    folderRoot = exportRoot & "\" & SafeName(accountSmtp) & NormalizeFolderPath(targetFolder.FolderPath)
    EnsureFolder folderRoot

    Set items = targetFolder.Items
    On Error Resume Next
    items.Sort "[ReceivedTime]", True
    On Error GoTo FolderError

    For itemIndex = 1 To items.Count
        Set currentItem = items(itemIndex)
        If TypeOf currentItem Is Outlook.MailItem Then
            Set mail = currentItem
            If Not exported.Exists(mail.EntryID) Then
                ExportMailItem mail, folderRoot
                exported.Add mail.EntryID, ""
                total = total + 1
            End If
        End If
    Next itemIndex

    For Each child In targetFolder.Folders
        total = total + ExportFolderTree(child, exportRoot, accountSmtp, exported)
    Next child

    ExportFolderTree = total
    Exit Function

FolderError:
    ExportFolderTree = total
End Function

Private Sub ExportMailItem(ByVal mail As Outlook.MailItem, ByVal folderRoot As String)
    On Error GoTo MailError

    Dim mailRoot As String
    Dim attachmentNames As Collection
    Dim metaPath As String

    mailRoot = folderRoot & "\" & BuildMailFolderName(mail)
    metaPath = mailRoot & "\meta.json"
    If FileExists(metaPath) Then Exit Sub

    EnsureFolder mailRoot

    mail.SaveAs mailRoot & "\mail.msg", OLMSGUNICODE
    WriteTextFile mailRoot & "\body.txt", mail.Body

    Set attachmentNames = SaveAttachments(mail, mailRoot)
    WriteMetaFile metaPath, mail, attachmentNames
    Exit Sub

MailError:
End Sub

Private Function SaveAttachments(ByVal mail As Outlook.MailItem, ByVal mailRoot As String) As Collection
    Dim result As New Collection
    Dim i As Long
    Dim item As Outlook.Attachment
    Dim safeFileName As String

    For i = 1 To mail.Attachments.Count
        Set item = mail.Attachments(i)
        safeFileName = SafeName(item.FileName)
        item.SaveAsFile mailRoot & "\" & safeFileName
        result.Add safeFileName
    Next i

    Set SaveAttachments = result
End Function

Private Sub WriteMetaFile(ByVal path As String, ByVal mail As Outlook.MailItem, ByVal attachmentNames As Collection)
    Dim body As String
    body = "{" & vbCrLf & _
        "  ""mail_id"": """ & JsonEscape(mail.EntryID) & """," & vbCrLf & _
        "  ""entry_id"": """ & JsonEscape(mail.EntryID) & """," & vbCrLf & _
        "  ""case_id"": """ & """," & vbCrLf & _
        "  ""sender_name"": """ & JsonEscape(mail.SenderName) & """," & vbCrLf & _
        "  ""sender_email"": """ & JsonEscape(GetSenderAddress(mail)) & """," & vbCrLf & _
        "  ""subject"": """ & JsonEscape(mail.Subject) & """," & vbCrLf & _
        "  ""received_at"": """ & Format$(mail.ReceivedTime, "yyyy-mm-dd\Thh:nn:ss") & """," & vbCrLf & _
        "  ""body_path"": ""body.txt""," & vbCrLf & _
        "  ""msg_path"": ""mail.msg""," & vbCrLf & _
        "  ""attachments"": " & CollectionToJsonArray(attachmentNames) & vbCrLf & _
        "}"

    WriteTextFile path, body
End Sub

' --- State file (exported EntryIDs) ---

Private Function LoadExportedIds(ByVal path As String) As Object
    Dim dict As Object
    Dim lineText As String
    Dim fileNumber As Integer
    Dim allText As String
    Dim entryId As String
    Dim pos As Long, startPos As Long

    Set dict = CreateObject("Scripting.Dictionary")

    If Dir$(path) = "" Then
        Set LoadExportedIds = dict
        Exit Function
    End If

    fileNumber = FreeFile
    Open path For Input As #fileNumber
    allText = ""
    Do Until EOF(fileNumber)
        Line Input #fileNumber, lineText
        allText = allText & lineText
    Loop
    Close #fileNumber

    ' Parse JSON array of strings: ["id1","id2",...]
    pos = 1
    Do
        pos = InStr(pos, allText, """")
        If pos = 0 Then Exit Do
        startPos = pos + 1
        pos = InStr(startPos, allText, """")
        If pos = 0 Then Exit Do
        entryId = Mid$(allText, startPos, pos - startPos)
        If Len(entryId) > 0 And Not dict.Exists(entryId) Then
            dict.Add entryId, ""
        End If
        pos = pos + 1
    Loop

    Set LoadExportedIds = dict
End Function

Private Sub SaveExportedIds(ByVal path As String, ByVal dict As Object)
    Dim fileNumber As Integer
    Dim key As Variant
    Dim first As Boolean

    EnsureFolder CreateObject("Scripting.FileSystemObject").GetParentFolderName(path)

    fileNumber = FreeFile
    Open path For Output As #fileNumber
    Print #fileNumber, "["
    first = True
    For Each key In dict.Keys
        If Not first Then Print #fileNumber, ","
        Print #fileNumber, "  """ & JsonEscape(CStr(key)) & """";
        first = False
    Next key
    Print #fileNumber, ""
    Print #fileNumber, "]"
    Close #fileNumber
End Sub

' --- Helpers ---

Private Function CollectionToJsonArray(ByVal values As Collection) As String
    Dim i As Long
    Dim text As String

    text = "["
    For i = 1 To values.Count
        If i > 1 Then text = text & ", "
        text = text & """" & JsonEscape(CStr(values(i))) & """"
    Next i
    text = text & "]"
    CollectionToJsonArray = text
End Function


Private Function GetStoreSmtpAddress(ByVal store As Outlook.Store) As String
    Dim account As Outlook.Account
    On Error Resume Next
    For Each account In Application.Session.Accounts
        If account.DeliveryStore.StoreID = store.StoreID Then
            GetStoreSmtpAddress = LCase$(account.SmtpAddress)
            Exit Function
        End If
    Next account
    GetStoreSmtpAddress = LCase$(store.GetRootFolder.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E"))
    On Error GoTo 0
End Function

Private Function GetSenderAddress(ByVal mail As Outlook.MailItem) As String
    On Error Resume Next
    GetSenderAddress = mail.SenderEmailAddress
    If Len(GetSenderAddress) = 0 Then
        GetSenderAddress = mail.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E")
    End If
    On Error GoTo 0
End Function

Private Function BuildMailFolderName(ByVal mail As Outlook.MailItem) As String
    BuildMailFolderName = Format$(mail.ReceivedTime, "yyyymmdd_hhnnss") & "_" & _
        SafeName(GetSenderAddress(mail)) & "_" & SafeName(mail.Subject) & "_" & Left$(SafeName(mail.EntryID), 40)
End Function

Private Function NormalizeFolderPath(ByVal folderPath As String) As String
    Dim parts() As String
    Dim i As Long
    Dim result As String

    parts = Split(folderPath, "\")
    result = ""
    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) > 0 Then
            result = result & "\" & SafeName(parts(i))
        End If
    Next i
    NormalizeFolderPath = result
End Function

Private Function SafeName(ByVal value As String) As String
    Dim text As String
    text = Trim$(value)
    If Len(text) = 0 Then text = "blank"
    text = Replace(text, "\", "_")
    text = Replace(text, "/", "_")
    text = Replace(text, ":", "_")
    text = Replace(text, "*", "_")
    text = Replace(text, "?", "_")
    text = Replace(text, Chr$(34), "_")
    text = Replace(text, "<", "_")
    text = Replace(text, ">", "_")
    text = Replace(text, "|", "_")
    If Len(text) > 80 Then text = Left$(text, 80)
    SafeName = text
End Function

Private Function JsonEscape(ByVal value As String) As String
    Dim text As String
    text = value
    text = Replace(text, "\", "\\")
    text = Replace(text, Chr$(34), "\"")
    text = Replace(text, vbCrLf, "\n")
    text = Replace(text, vbCr, "\n")
    text = Replace(text, vbLf, "\n")
    JsonEscape = text
End Function

Private Sub WriteTextFile(ByVal path As String, ByVal contents As String)
    Dim fileNumber As Integer
    fileNumber = FreeFile
    Open path For Output As #fileNumber
    Print #fileNumber, contents
    Close #fileNumber
End Sub

Private Sub EnsureFolder(ByVal path As String)
    Dim fso As Object
    Dim parentPath As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(path) Then Exit Sub

    parentPath = fso.GetParentFolderName(path)
    If Len(parentPath) > 0 Then
        If Not fso.FolderExists(parentPath) Then
            EnsureFolder parentPath
        End If
    End If

    If Not fso.FolderExists(path) Then
        fso.CreateFolder path
    End If
End Sub

Private Function FileExists(ByVal path As String) As Boolean
    FileExists = (Dir$(path) <> "")
End Function
