Attribute VB_Name = "AM_CRM_Outlook_Helper"
Option Explicit

' AM_CRM Outlook helper.
' Import this module into an Outlook or Excel VBA project.
' These routines only display drafts or export logs. They never send mail.

' Outlook MailItem type used by late-bound CreateItem.
Private Const OL_MAIL_ITEM As Long = 0

' Outlook Inbox folder id used by late-bound GetDefaultFolder.
Private Const OL_FOLDER_INBOX As Long = 6

' Outlook Sent Mail folder id used by late-bound GetDefaultFolder.
Private Const OL_FOLDER_SENT_MAIL As Long = 5

' Default number of Outlook messages exported per client email address.
Private Const DEFAULT_MAX_EMAILS_PER_ADDRESS As Long = 20

' Default number of calendar days scanned backward for matching Outlook emails.
Private Const DEFAULT_LOOKBACK_DAYS As Long = 365

' Default number of email body characters written to JSON previews.
Private Const DEFAULT_BODY_PREVIEW_CHARS As Long = 750

' Default output filename for recent inbox and sent email history.
Private Const DEFAULT_RECENT_EMAIL_OUTPUT_FILE As String = "AM_CRM_Outlook_Recent_Emails.json"

' Default output filename for sent email contact logging.
Private Const DEFAULT_SENT_LOG_OUTPUT_FILE As String = "AM_CRM_Outlook_Sent_Email_Log.json"

' Default output filename for Outlook email task candidates.
Private Const DEFAULT_TASK_OUTPUT_FILE As String = "AM_CRM_Outlook_Task_Candidates.json"

' Default task title prefix used for email-to-task exports.
Private Const DEFAULT_TASK_TITLE_PREFIX As String = "Follow up: "

' Default priority assigned to exported email task candidates unless Outlook marks the email important.
Private Const DEFAULT_TASK_PRIORITY As String = "Normal"

' Default status assigned to exported email task candidates.
Private Const DEFAULT_TASK_STATUS As String = "Open"

' MAPI property tag used to resolve Exchange address entries to SMTP addresses.
Private Const PR_SMTP_ADDRESS As String = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"

Public Sub AMCRM_OpenDraftEmail(ByVal toList As String, Optional ByVal subjectText As String = "", Optional ByVal bodyText As String = "")
    Dim olApp As Object
    Dim mail As Object

    Set olApp = GetOutlookApplication()
    If olApp Is Nothing Then
        MsgBox "Outlook is not available.", vbExclamation, "AM_CRM"
        Exit Sub
    End If

    Set mail = olApp.CreateItem(OL_MAIL_ITEM)
    With mail
        .To = toList
        .Subject = subjectText
        .Body = bodyText
        .Display
    End With
End Sub

Public Sub AMCRM_OpenDraftFromJsonFile(Optional ByVal draftJsonPath As String = "")
    Dim jsonText As String
    Dim recipients As Collection
    Dim toList As String
    Dim subjectText As String
    Dim bodyText As String

    If Len(Trim$(draftJsonPath)) = 0 Then
        draftJsonPath = PromptForPath("Draft JSON path", "AM_CRM_Outlook_Draft.json")
    End If
    If Len(Trim$(draftJsonPath)) = 0 Then Exit Sub

    jsonText = ReadTextFile(draftJsonPath)
    Set recipients = JsonArrayStrings(jsonText, "to")
    toList = JoinCollection(recipients, "; ")
    If Len(toList) = 0 Then toList = JsonStringValue(jsonText, "to")

    subjectText = JsonStringValue(jsonText, "subject")
    bodyText = JsonStringValue(jsonText, "body")
    If Len(toList) = 0 Then
        MsgBox "Draft JSON does not contain recipients.", vbExclamation, "AM_CRM"
        Exit Sub
    End If

    AMCRM_OpenDraftEmail toList, subjectText, bodyText
End Sub

Public Sub AMCRM_ExportRecentEmails(ByVal clientEmailList As String, ByVal outputJsonPath As String, Optional ByVal maxPerAddress As Long = -1, Optional ByVal lookbackDays As Long = -1, Optional ByVal bodyPreviewChars As Long = -1)
    ExportRecentEmailContacts ParseEmailListToContacts(clientEmailList), outputJsonPath, maxPerAddress, lookbackDays, bodyPreviewChars
End Sub

Public Sub AMCRM_ExportRecentEmailsFromClientJson(Optional ByVal clientJsonPath As String = "", Optional ByVal outputJsonPath As String = "", Optional ByVal maxPerAddress As Long = -1, Optional ByVal lookbackDays As Long = -1, Optional ByVal bodyPreviewChars As Long = -1)
    Dim contacts As Collection

    Set contacts = ContactsFromClientJsonPath(clientJsonPath)
    If contacts Is Nothing Then Exit Sub

    If Len(Trim$(outputJsonPath)) = 0 Then
        outputJsonPath = PromptForPath("Recent email output JSON path", DEFAULT_RECENT_EMAIL_OUTPUT_FILE)
    End If
    If Len(Trim$(outputJsonPath)) = 0 Then Exit Sub

    ExportRecentEmailContacts contacts, outputJsonPath, maxPerAddress, lookbackDays, bodyPreviewChars
End Sub

Public Sub AMCRM_ExportSentEmailLog(ByVal clientEmailList As String, ByVal outputJsonPath As String, Optional ByVal maxPerAddress As Long = -1, Optional ByVal lookbackDays As Long = -1, Optional ByVal bodyPreviewChars As Long = -1)
    ExportSentEmailContacts ParseEmailListToContacts(clientEmailList), outputJsonPath, maxPerAddress, lookbackDays, bodyPreviewChars
End Sub

Public Sub AMCRM_ExportSentEmailLogFromClientJson(Optional ByVal clientJsonPath As String = "", Optional ByVal outputJsonPath As String = "", Optional ByVal maxPerAddress As Long = -1, Optional ByVal lookbackDays As Long = -1, Optional ByVal bodyPreviewChars As Long = -1)
    Dim contacts As Collection

    Set contacts = ContactsFromClientJsonPath(clientJsonPath)
    If contacts Is Nothing Then Exit Sub

    If Len(Trim$(outputJsonPath)) = 0 Then
        outputJsonPath = PromptForPath("Sent email log output JSON path", DEFAULT_SENT_LOG_OUTPUT_FILE)
    End If
    If Len(Trim$(outputJsonPath)) = 0 Then Exit Sub

    ExportSentEmailContacts contacts, outputJsonPath, maxPerAddress, lookbackDays, bodyPreviewChars
End Sub

Public Sub AMCRM_ExportTaskCandidatesFromClientJson(Optional ByVal clientJsonPath As String = "", Optional ByVal outputJsonPath As String = "", Optional ByVal lookbackDays As Long = -1, Optional ByVal bodyPreviewChars As Long = -1, Optional ByVal includeUnreadInbox As Boolean = True, Optional ByVal includeFlagged As Boolean = True)
    Dim olApp As Object
    Dim ns As Object
    Dim contacts As Collection
    Dim tasksJson As String
    Dim json As String

    Set contacts = ContactsFromClientJsonPath(clientJsonPath)
    If contacts Is Nothing Then Exit Sub

    If Len(Trim$(outputJsonPath)) = 0 Then
        outputJsonPath = PromptForPath("Task candidate output JSON path", DEFAULT_TASK_OUTPUT_FILE)
    End If
    If Len(Trim$(outputJsonPath)) = 0 Then Exit Sub

    Set olApp = GetOutlookApplication()
    If olApp Is Nothing Then
        MsgBox "Outlook is not available.", vbExclamation, "AM_CRM"
        Exit Sub
    End If

    Set ns = olApp.GetNamespace("MAPI")
    lookbackDays = PositiveOrDefault(lookbackDays, DEFAULT_LOOKBACK_DAYS)
    bodyPreviewChars = NonNegativeOrDefault(bodyPreviewChars, DEFAULT_BODY_PREVIEW_CHARS)

    tasksJson = TrimTrailingComma(ExportTaskFolderMatches(ns.GetDefaultFolder(OL_FOLDER_INBOX), contacts, lookbackDays, bodyPreviewChars, includeUnreadInbox, includeFlagged))
    json = "{""exportType"":""AM_CRM_OUTLOOK_TASK_CANDIDATES"",""generatedAt"":""" & JsonEscape(FormatIso(Now)) & """,""source"":""AM_CRM_Outlook_Helper"",""tasks"":["
    json = json & tasksJson
    json = json & "]}"

    WriteTextFile outputJsonPath, json
    MsgBox "Outlook task candidate export complete.", vbInformation, "AM_CRM"
End Sub

Public Sub AMCRM_OpenDraftPrompt()
    Dim toList As String
    Dim subjectText As String
    Dim bodyText As String

    toList = InputBox("To", "AM_CRM Outlook Draft")
    If Len(Trim$(toList)) = 0 Then Exit Sub
    subjectText = InputBox("Subject", "AM_CRM Outlook Draft")
    bodyText = InputBox("Body", "AM_CRM Outlook Draft")
    AMCRM_OpenDraftEmail toList, subjectText, bodyText
End Sub

Private Sub ExportRecentEmailContacts(ByVal contacts As Collection, ByVal outputJsonPath As String, ByVal maxPerAddress As Long, ByVal lookbackDays As Long, ByVal bodyPreviewChars As Long)
    Dim olApp As Object
    Dim ns As Object
    Dim inboxJson As String
    Dim sentJson As String
    Dim json As String

    If contacts Is Nothing Then
        MsgBox "No client email addresses were provided.", vbExclamation, "AM_CRM"
        Exit Sub
    End If
    If contacts.Count = 0 Then
        MsgBox "No client email addresses were provided.", vbExclamation, "AM_CRM"
        Exit Sub
    End If

    If Len(Trim$(outputJsonPath)) = 0 Then
        outputJsonPath = PromptForPath("Recent email output JSON path", DEFAULT_RECENT_EMAIL_OUTPUT_FILE)
    End If
    If Len(Trim$(outputJsonPath)) = 0 Then Exit Sub

    Set olApp = GetOutlookApplication()
    If olApp Is Nothing Then
        MsgBox "Outlook is not available.", vbExclamation, "AM_CRM"
        Exit Sub
    End If

    Set ns = olApp.GetNamespace("MAPI")
    maxPerAddress = PositiveOrDefault(maxPerAddress, DEFAULT_MAX_EMAILS_PER_ADDRESS)
    lookbackDays = PositiveOrDefault(lookbackDays, DEFAULT_LOOKBACK_DAYS)
    bodyPreviewChars = NonNegativeOrDefault(bodyPreviewChars, DEFAULT_BODY_PREVIEW_CHARS)

    inboxJson = TrimTrailingComma(ExportFolderMatches(ns.GetDefaultFolder(OL_FOLDER_INBOX), contacts, maxPerAddress, "Inbox", lookbackDays, bodyPreviewChars))
    sentJson = TrimTrailingComma(ExportFolderMatches(ns.GetDefaultFolder(OL_FOLDER_SENT_MAIL), contacts, maxPerAddress, "Sent", lookbackDays, bodyPreviewChars))

    json = "{""exportType"":""AM_CRM_OUTLOOK_RECENT_EMAILS"",""generatedAt"":""" & JsonEscape(FormatIso(Now)) & """,""source"":""AM_CRM_Outlook_Helper"",""emails"":["
    json = json & inboxJson
    If Len(inboxJson) > 0 And Len(sentJson) > 0 Then json = json & ","
    json = json & sentJson
    json = json & "]}"

    WriteTextFile outputJsonPath, json
    MsgBox "Recent email export complete.", vbInformation, "AM_CRM"
End Sub

Private Sub ExportSentEmailContacts(ByVal contacts As Collection, ByVal outputJsonPath As String, ByVal maxPerAddress As Long, ByVal lookbackDays As Long, ByVal bodyPreviewChars As Long)
    Dim olApp As Object
    Dim ns As Object
    Dim sentJson As String
    Dim json As String

    If contacts Is Nothing Then
        MsgBox "No client email addresses were provided.", vbExclamation, "AM_CRM"
        Exit Sub
    End If
    If contacts.Count = 0 Then
        MsgBox "No client email addresses were provided.", vbExclamation, "AM_CRM"
        Exit Sub
    End If

    If Len(Trim$(outputJsonPath)) = 0 Then
        outputJsonPath = PromptForPath("Sent email log output JSON path", DEFAULT_SENT_LOG_OUTPUT_FILE)
    End If
    If Len(Trim$(outputJsonPath)) = 0 Then Exit Sub

    Set olApp = GetOutlookApplication()
    If olApp Is Nothing Then
        MsgBox "Outlook is not available.", vbExclamation, "AM_CRM"
        Exit Sub
    End If

    Set ns = olApp.GetNamespace("MAPI")
    maxPerAddress = PositiveOrDefault(maxPerAddress, DEFAULT_MAX_EMAILS_PER_ADDRESS)
    lookbackDays = PositiveOrDefault(lookbackDays, DEFAULT_LOOKBACK_DAYS)
    bodyPreviewChars = NonNegativeOrDefault(bodyPreviewChars, DEFAULT_BODY_PREVIEW_CHARS)

    sentJson = TrimTrailingComma(ExportFolderMatches(ns.GetDefaultFolder(OL_FOLDER_SENT_MAIL), contacts, maxPerAddress, "Sent", lookbackDays, bodyPreviewChars))
    json = "{""exportType"":""AM_CRM_OUTLOOK_SENT_EMAIL_LOG"",""generatedAt"":""" & JsonEscape(FormatIso(Now)) & """,""source"":""AM_CRM_Outlook_Helper"",""sentEmails"":["
    json = json & sentJson
    json = json & "]}"

    WriteTextFile outputJsonPath, json
    MsgBox "Sent email log export complete.", vbInformation, "AM_CRM"
End Sub

Private Function ExportFolderMatches(ByVal folder As Object, ByVal contacts As Collection, ByVal maxPerAddress As Long, ByVal folderLabel As String, ByVal lookbackDays As Long, ByVal bodyPreviewChars As Long) As String
    Dim items As Object
    Dim item As Object
    Dim contact As Object
    Dim counts As Object
    Dim json As String
    Dim emailKey As String
    Dim whenValue As Date
    Dim cutoff As Date

    Set counts = CreateObject("Scripting.Dictionary")
    InitializeContactCounts counts, contacts

    Set items = folder.Items
    items.Sort "[" & IIf(folderLabel = "Sent", "SentOn", "ReceivedTime") & "]", True
    cutoff = DateAdd("d", -lookbackDays, Now)

    json = ""
    For Each item In items
        If TypeName(item) = "MailItem" Then
            whenValue = MailWhen(item, folderLabel)
            If lookbackDays > 0 And whenValue < cutoff Then Exit For

            Set contact = ClientContactForMail(item, contacts)
            If Not contact Is Nothing Then
                emailKey = NormalizeEmail(DictValue(contact, "email"))
                If counts(emailKey) < maxPerAddress Then
                    counts(emailKey) = counts(emailKey) + 1
                    json = json & MailItemJson(item, contact, folderLabel, bodyPreviewChars) & ","
                End If
            End If
        End If
    Next item

    ExportFolderMatches = json
End Function

Private Function ExportTaskFolderMatches(ByVal folder As Object, ByVal contacts As Collection, ByVal lookbackDays As Long, ByVal bodyPreviewChars As Long, ByVal includeUnreadInbox As Boolean, ByVal includeFlagged As Boolean) As String
    Dim items As Object
    Dim item As Object
    Dim contact As Object
    Dim json As String
    Dim cutoff As Date
    Dim whenValue As Date

    Set items = folder.Items
    items.Sort "[ReceivedTime]", True
    cutoff = DateAdd("d", -lookbackDays, Now)

    json = ""
    For Each item In items
        If TypeName(item) = "MailItem" Then
            whenValue = MailWhen(item, "Inbox")
            If lookbackDays > 0 And whenValue < cutoff Then Exit For
            If ShouldExportTaskCandidate(item, includeUnreadInbox, includeFlagged) Then
                Set contact = ClientContactForMail(item, contacts)
                If Not contact Is Nothing Then
                    json = json & MailItemTaskJson(item, contact, bodyPreviewChars) & ","
                End If
            End If
        End If
    Next item

    ExportTaskFolderMatches = json
End Function

Private Function ShouldExportTaskCandidate(ByVal mail As Object, ByVal includeUnreadInbox As Boolean, ByVal includeFlagged As Boolean) As Boolean
    Dim isUnread As Boolean
    Dim isFlagged As Boolean

    On Error Resume Next
    isUnread = CBool(mail.UnRead)
    isFlagged = (CLng(mail.FlagStatus) <> 0) Or Len(Trim$(mail.FlagRequest & "")) > 0
    On Error GoTo 0

    ShouldExportTaskCandidate = (includeUnreadInbox And isUnread) Or (includeFlagged And isFlagged)
End Function

Private Function ClientContactForMail(ByVal mail As Object, ByVal contacts As Collection) As Object
    Dim contact As Variant
    Dim emailKey As String
    Dim participantText As String

    participantText = LCase$(MailParticipantText(mail))
    For Each contact In contacts
        emailKey = NormalizeEmail(DictValue(contact, "email"))
        If Len(emailKey) > 0 And InStr(1, participantText, emailKey, vbTextCompare) > 0 Then
            Set ClientContactForMail = contact
            Exit Function
        End If
    Next contact
End Function

Private Function MailParticipantText(ByVal mail As Object) As String
    MailParticipantText = LCase$(MailSenderEmailAddress(mail) & " " & mail.SenderName & " " & RecipientAddresses(mail))
End Function

Private Function RecipientAddresses(ByVal mail As Object) As String
    Dim recipient As Object
    Dim result As String
    Dim smtpAddress As String

    result = ""
    For Each recipient In mail.Recipients
        smtpAddress = ""
        On Error Resume Next
        smtpAddress = SmtpAddressFromAddressEntry(recipient.AddressEntry)
        If Len(smtpAddress) = 0 Then smtpAddress = recipient.Address
        result = result & " " & smtpAddress & " " & recipient.Name
        On Error GoTo 0
    Next recipient
    RecipientAddresses = result
End Function

Private Function MailSenderEmailAddress(ByVal mail As Object) As String
    Dim smtpAddress As String

    smtpAddress = ""
    On Error Resume Next
    If Not mail.Sender Is Nothing Then smtpAddress = SmtpAddressFromAddressEntry(mail.Sender)
    If Len(smtpAddress) = 0 Then smtpAddress = mail.SenderEmailAddress
    On Error GoTo 0

    MailSenderEmailAddress = smtpAddress
End Function

Private Function SmtpAddressFromAddressEntry(ByVal addressEntry As Object) As String
    Dim exchangeUser As Object
    Dim smtpAddress As String

    smtpAddress = ""
    On Error Resume Next
    Set exchangeUser = addressEntry.GetExchangeUser
    If Not exchangeUser Is Nothing Then smtpAddress = exchangeUser.PrimarySmtpAddress
    If Len(smtpAddress) = 0 Then smtpAddress = addressEntry.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS)
    On Error GoTo 0

    SmtpAddressFromAddressEntry = smtpAddress
End Function

Private Function MailItemJson(ByVal mail As Object, ByVal contact As Object, ByVal folderLabel As String, ByVal bodyPreviewChars As Long) As String
    MailItemJson = "{""clientGroupId"":""" & JsonEscape(DictValue(contact, "clientGroupId")) & """," & _
        """clientName"":""" & JsonEscape(DictValue(contact, "clientName")) & """," & _
        """matchedAddress"":""" & JsonEscape(DictValue(contact, "email")) & """," & _
        """folder"":""" & JsonEscape(folderLabel) & """," & _
        """timestamp"":""" & JsonEscape(FormatIso(MailWhen(mail, folderLabel))) & """," & _
        """outlookEntryId"":""" & JsonEscape(mail.EntryID) & """," & _
        """sender"":""" & JsonEscape(mail.SenderName) & """," & _
        """senderEmail"":""" & JsonEscape(MailSenderEmailAddress(mail)) & """," & _
        """to"":""" & JsonEscape(RecipientAddresses(mail)) & """," & _
        """subject"":""" & JsonEscape(mail.Subject) & """," & _
        """bodyPreview"":""" & JsonEscape(Left$(mail.Body, bodyPreviewChars)) & """}"
End Function

Private Function MailItemTaskJson(ByVal mail As Object, ByVal contact As Object, ByVal bodyPreviewChars As Long) As String
    Dim title As String
    Dim notes As String
    Dim priority As String
    Dim dueDate As String

    title = DEFAULT_TASK_TITLE_PREFIX & Trim$(mail.Subject & "")
    If title = DEFAULT_TASK_TITLE_PREFIX Then title = DEFAULT_TASK_TITLE_PREFIX & "(no subject)"
    notes = "Outlook email from " & mail.SenderName & " on " & FormatIso(MailWhen(mail, "Inbox")) & vbCrLf & Left$(mail.Body, bodyPreviewChars)
    priority = DEFAULT_TASK_PRIORITY

    On Error Resume Next
    If CLng(mail.Importance) >= 2 Then priority = "High"
    If Year(mail.TaskDueDate) > 1900 Then dueDate = Format$(mail.TaskDueDate, "yyyy-mm-dd")
    On Error GoTo 0

    MailItemTaskJson = "{""clientGroupId"":""" & JsonEscape(DictValue(contact, "clientGroupId")) & """," & _
        """clientName"":""" & JsonEscape(DictValue(contact, "clientName")) & """," & _
        """matchedAddress"":""" & JsonEscape(DictValue(contact, "email")) & """," & _
        """title"":""" & JsonEscape(title) & """," & _
        """dueDate"":""" & JsonEscape(dueDate) & """," & _
        """priority"":""" & JsonEscape(priority) & """," & _
        """status"":""" & JsonEscape(DEFAULT_TASK_STATUS) & """," & _
        """notes"":""" & JsonEscape(notes) & """," & _
        """source"":""outlook_vba_helper""," & _
        """sourceColumn"":""tasks""," & _
        """outlookEntryId"":""" & JsonEscape(mail.EntryID) & """," & _
        """emailTimestamp"":""" & JsonEscape(FormatIso(MailWhen(mail, "Inbox"))) & """," & _
        """subject"":""" & JsonEscape(mail.Subject) & """}"
End Function

Private Function MailWhen(ByVal mail As Object, ByVal folderLabel As String) As Date
    On Error Resume Next
    If folderLabel = "Sent" Then
        MailWhen = mail.SentOn
    Else
        MailWhen = mail.ReceivedTime
    End If
    If Err.Number <> 0 Then
        Err.Clear
        MailWhen = mail.CreationTime
    End If
    On Error GoTo 0
End Function

Private Function ContactsFromClientJsonPath(ByVal clientJsonPath As String) As Collection
    Dim jsonText As String
    Dim contacts As Collection

    If Len(Trim$(clientJsonPath)) = 0 Then
        clientJsonPath = PromptForPath("Client email JSON path", "AM_CRM_Outlook_Client_Emails.json")
    End If
    If Len(Trim$(clientJsonPath)) = 0 Then Exit Function

    jsonText = ReadTextFile(clientJsonPath)
    Set contacts = ParseClientContactsFromJson(jsonText)
    If contacts.Count = 0 Then
        MsgBox "No client contacts were found in the JSON file.", vbExclamation, "AM_CRM"
        Exit Function
    End If

    Set ContactsFromClientJsonPath = contacts
End Function

Private Function ParseClientContactsFromJson(ByVal jsonText As String) As Collection
    Dim contacts As New Collection
    Dim arrayText As String
    Dim chunks As Collection
    Dim chunk As Variant
    Dim clientGroupId As String
    Dim clientName As String
    Dim email As String

    arrayText = JsonArraySection(jsonText, "contacts")
    If Len(arrayText) > 0 Then
        Set chunks = JsonObjectChunks(arrayText)
        For Each chunk In chunks
            clientGroupId = JsonStringValue(CStr(chunk), "clientGroupId")
            clientName = JsonStringValue(CStr(chunk), "clientName")
            email = JsonStringValue(CStr(chunk), "email")
            If Len(Trim$(email)) > 0 Then contacts.Add ContactDictionary(email, clientGroupId, clientName)
        Next chunk
    End If

    If contacts.Count = 0 Then
        Set contacts = ParseEmailListToContacts(jsonText)
    End If
    Set ParseClientContactsFromJson = contacts
End Function

Private Function ParseEmailListToContacts(ByVal clientEmailList As String) As Collection
    Dim result As New Collection
    Dim normalized As String
    Dim parts() As String
    Dim i As Long
    Dim value As String

    normalized = Replace(clientEmailList, vbCr, ";")
    normalized = Replace(normalized, vbLf, ";")
    normalized = Replace(normalized, ",", ";")
    parts = Split(normalized, ";")

    For i = LBound(parts) To UBound(parts)
        value = Trim$(parts(i))
        If Len(value) > 0 And InStr(1, value, "@", vbTextCompare) > 0 Then
            result.Add ContactDictionary(value, "", "")
        End If
    Next i

    Set ParseEmailListToContacts = result
End Function

Private Function ContactDictionary(ByVal email As String, ByVal clientGroupId As String, ByVal clientName As String) As Object
    Dim contact As Object

    Set contact = CreateObject("Scripting.Dictionary")
    contact("email") = Trim$(email)
    contact("clientGroupId") = Trim$(clientGroupId)
    contact("clientName") = Trim$(clientName)
    Set ContactDictionary = contact
End Function

Private Sub InitializeContactCounts(ByVal counts As Object, ByVal contacts As Collection)
    Dim contact As Variant
    Dim emailKey As String

    For Each contact In contacts
        emailKey = NormalizeEmail(DictValue(contact, "email"))
        If Len(emailKey) > 0 And Not counts.Exists(emailKey) Then counts(emailKey) = 0
    Next contact
End Sub

Private Function DictValue(ByVal dict As Object, ByVal key As String) As String
    If dict.Exists(key) Then DictValue = CStr(dict(key)) Else DictValue = ""
End Function

Private Function NormalizeEmail(ByVal value As String) As String
    NormalizeEmail = LCase$(Trim$(value))
End Function

Private Function PositiveOrDefault(ByVal value As Long, ByVal fallback As Long) As Long
    If value > 0 Then PositiveOrDefault = value Else PositiveOrDefault = fallback
End Function

Private Function NonNegativeOrDefault(ByVal value As Long, ByVal fallback As Long) As Long
    If value >= 0 Then NonNegativeOrDefault = value Else NonNegativeOrDefault = fallback
End Function

Private Function TrimTrailingComma(ByVal value As String) As String
    If Len(value) > 0 And Right$(value, 1) = "," Then
        TrimTrailingComma = Left$(value, Len(value) - 1)
    Else
        TrimTrailingComma = value
    End If
End Function

Private Function FormatIso(ByVal value As Date) As String
    FormatIso = Format$(value, "yyyy-mm-dd\Thh:nn:ss")
End Function

Private Function PromptForPath(ByVal promptText As String, ByVal defaultFileName As String) As String
    PromptForPath = Trim$(InputBox(promptText, "AM_CRM Outlook Helper", DefaultDownloadPath(defaultFileName)))
End Function

Private Function DefaultDownloadPath(ByVal fileName As String) As String
    Dim homePath As String

    homePath = Environ$("USERPROFILE")
    If Len(homePath) > 0 Then
        DefaultDownloadPath = homePath & "\Downloads\" & fileName
    Else
        DefaultDownloadPath = CurDir$ & "\" & fileName
    End If
End Function

Private Function GetOutlookApplication() As Object
    On Error Resume Next
    Set GetOutlookApplication = GetObject(, "Outlook.Application")
    If GetOutlookApplication Is Nothing Then
        Set GetOutlookApplication = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0
End Function

Private Function ReadTextFile(ByVal inputPath As String) As String
    Dim stream As Object
    Dim fso As Object

    On Error GoTo Fallback
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile inputPath
    ReadTextFile = stream.ReadText
    stream.Close
    Exit Function

Fallback:
    On Error GoTo 0
    Set fso = CreateObject("Scripting.FileSystemObject")
    ReadTextFile = fso.OpenTextFile(inputPath, 1, False).ReadAll
End Function

Private Sub WriteTextFile(ByVal outputPath As String, ByVal content As String)
    Dim stream As Object
    Dim fso As Object
    Dim fallbackStream As Object

    On Error GoTo Fallback
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText content
    stream.SaveToFile outputPath, 2
    stream.Close
    Exit Sub

Fallback:
    On Error GoTo 0
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fallbackStream = fso.CreateTextFile(outputPath, True, True)
    fallbackStream.Write content
    fallbackStream.Close
End Sub

Private Function JsonStringValue(ByVal jsonText As String, ByVal key As String) As String
    Dim marker As String
    Dim pos As Long
    Dim colonPos As Long
    Dim quotePos As Long

    marker = """" & key & """"
    pos = InStr(1, jsonText, marker, vbTextCompare)
    If pos = 0 Then Exit Function

    colonPos = InStr(pos + Len(marker), jsonText, ":")
    If colonPos = 0 Then Exit Function

    quotePos = InStr(colonPos + 1, jsonText, """")
    If quotePos = 0 Then Exit Function

    JsonStringValue = ReadJsonQuotedString(jsonText, quotePos)
End Function

Private Function JsonArrayStrings(ByVal jsonText As String, ByVal key As String) As Collection
    Dim result As New Collection
    Dim arrayText As String
    Dim i As Long
    Dim ch As String

    arrayText = JsonArraySection(jsonText, key)
    i = 1
    Do While i <= Len(arrayText)
        ch = Mid$(arrayText, i, 1)
        If ch = """" Then
            result.Add ReadJsonQuotedString(arrayText, i)
            i = JsonStringEndPosition(arrayText, i)
        End If
        i = i + 1
    Loop

    Set JsonArrayStrings = result
End Function

Private Function JsonArraySection(ByVal jsonText As String, ByVal key As String) As String
    Dim marker As String
    Dim pos As Long
    Dim openPos As Long
    Dim i As Long
    Dim depth As Long
    Dim ch As String
    Dim inString As Boolean
    Dim escaped As Boolean

    marker = """" & key & """"
    pos = InStr(1, jsonText, marker, vbTextCompare)
    If pos = 0 Then Exit Function

    openPos = InStr(pos + Len(marker), jsonText, "[")
    If openPos = 0 Then Exit Function

    depth = 1
    For i = openPos + 1 To Len(jsonText)
        ch = Mid$(jsonText, i, 1)
        If inString Then
            If escaped Then
                escaped = False
            ElseIf ch = "\" Then
                escaped = True
            ElseIf ch = """" Then
                inString = False
            End If
        Else
            If ch = """" Then
                inString = True
            ElseIf ch = "[" Then
                depth = depth + 1
            ElseIf ch = "]" Then
                depth = depth - 1
                If depth = 0 Then
                    JsonArraySection = Mid$(jsonText, openPos + 1, i - openPos - 1)
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

Private Function JsonObjectChunks(ByVal arrayText As String) As Collection
    Dim result As New Collection
    Dim i As Long
    Dim depth As Long
    Dim startPos As Long
    Dim ch As String
    Dim inString As Boolean
    Dim escaped As Boolean

    For i = 1 To Len(arrayText)
        ch = Mid$(arrayText, i, 1)
        If inString Then
            If escaped Then
                escaped = False
            ElseIf ch = "\" Then
                escaped = True
            ElseIf ch = """" Then
                inString = False
            End If
        Else
            If ch = """" Then
                inString = True
            ElseIf ch = "{" Then
                If depth = 0 Then startPos = i
                depth = depth + 1
            ElseIf ch = "}" Then
                depth = depth - 1
                If depth = 0 And startPos > 0 Then result.Add Mid$(arrayText, startPos, i - startPos + 1)
            End If
        End If
    Next i

    Set JsonObjectChunks = result
End Function

Private Function ReadJsonQuotedString(ByVal jsonText As String, ByVal quotePos As Long) As String
    Dim i As Long
    Dim ch As String
    Dim escaped As Boolean
    Dim result As String
    Dim code As String

    For i = quotePos + 1 To Len(jsonText)
        ch = Mid$(jsonText, i, 1)
        If escaped Then
            Select Case ch
                Case """"
                    result = result & """"
                Case "\"
                    result = result & "\"
                Case "/"
                    result = result & "/"
                Case "b"
                    result = result & Chr$(8)
                Case "f"
                    result = result & Chr$(12)
                Case "n"
                    result = result & vbLf
                Case "r"
                    result = result & vbCr
                Case "t"
                    result = result & vbTab
                Case "u"
                    code = Mid$(jsonText, i + 1, 4)
                    If IsHex4(code) Then
                        result = result & ChrW$(CLng("&H" & code))
                        i = i + 4
                    Else
                        result = result & ch
                    End If
                Case Else
                    result = result & ch
            End Select
            escaped = False
        ElseIf ch = "\" Then
            escaped = True
        ElseIf ch = """" Then
            Exit For
        Else
            result = result & ch
        End If
    Next i

    ReadJsonQuotedString = result
End Function

Private Function JsonStringEndPosition(ByVal jsonText As String, ByVal quotePos As Long) As Long
    Dim i As Long
    Dim ch As String
    Dim escaped As Boolean

    For i = quotePos + 1 To Len(jsonText)
        ch = Mid$(jsonText, i, 1)
        If escaped Then
            escaped = False
        ElseIf ch = "\" Then
            escaped = True
        ElseIf ch = """" Then
            JsonStringEndPosition = i
            Exit Function
        End If
    Next i

    JsonStringEndPosition = Len(jsonText)
End Function

Private Function IsHex4(ByVal value As String) As Boolean
    IsHex4 = Len(value) = 4 And UCase$(value) Like "[0-9A-F][0-9A-F][0-9A-F][0-9A-F]"
End Function

Private Function JoinCollection(ByVal values As Collection, ByVal separator As String) As String
    Dim value As Variant
    Dim result As String

    For Each value In values
        If Len(result) > 0 Then result = result & separator
        result = result & CStr(value)
    Next value
    JoinCollection = result
End Function

Private Function JsonEscape(ByVal value As String) As String
    Dim s As String

    s = value
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    JsonEscape = s
End Function
