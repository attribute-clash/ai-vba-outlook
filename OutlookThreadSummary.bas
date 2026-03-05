Attribute VB_Name = "OutlookThreadSummary"
Option Explicit

Private Const ENV_API_KEY As String = "MY_API_KEY"
Private Const ENV_API_URL As String = "LLM_API_URL"
Private Const ENV_MODEL As String = "LLM_MODEL"
Private Const DEFAULT_API_URL As String = "https://api.openai.com/v1/chat/completions"
Private Const DEFAULT_MODEL As String = "gpt-4o-mini"
Private Const REQUEST_TIMEOUT_MS As Long = 300000
Private Const MAX_CONTEXT_CHARS As Long = 120000

' Кнопка на ленте может вызывать этот метод как onAction="RunThreadSummary"
Public Sub RunThreadSummary(Optional control As IRibbonControl)
    On Error GoTo FatalError

    Dim sourceMail As Outlook.MailItem
    Set sourceMail = GetCurrentMailItem()

    If sourceMail Is Nothing Then
        MsgBox "Пожалуйста, откройте письмо или выберите его в списке", vbExclamation
        Exit Sub
    End If

    Dim mails As Collection
    Set mails = BuildConversationMailCollection(sourceMail)

    If mails Is Nothing Or mails.Count = 0 Then
        MsgBox "Не удалось собрать письма из текущей цепочки.", vbExclamation
        Exit Sub
    End If

    Dim timeline As Collection
    Set timeline = SortMailsByDateAsc(mails)

    Dim payloadText As String
    payloadText = BuildUserPrompt(timeline)

    If Len(payloadText) = 0 Then
        MsgBox "Цепочка пуста после очистки текста.", vbExclamation
        Exit Sub
    End If

    Dim truncWarn As String
    payloadText = EnforceContextLimit(payloadText, truncWarn)

    Dim summary As String
    summary = RequestSummaryFromLlm(payloadText)

    If Len(summary) = 0 Then
        MsgBox "API вернул пустой ответ.", vbExclamation
        Exit Sub
    End If

    Dim outPath As String
    outPath = SaveSummaryToDesktop(summary)

    Dim doneMessage As String
    doneMessage = "Саммари сохранено в файл:" & vbCrLf & outPath
    If Len(truncWarn) > 0 Then doneMessage = doneMessage & vbCrLf & vbCrLf & truncWarn

    MsgBox doneMessage, vbInformation
    Exit Sub

FatalError:
    MsgBox "Ошибка макроса: " & Err.Description, vbCritical
End Sub

Private Function GetCurrentMailItem() As Outlook.MailItem
    On Error Resume Next

    Dim inspector As Outlook.Inspector
    Set inspector = Application.ActiveInspector
    If Not inspector Is Nothing Then
        If TypeOf inspector.CurrentItem Is Outlook.MailItem Then
            Set GetCurrentMailItem = inspector.CurrentItem
            Exit Function
        End If
    End If

    Dim explorer As Outlook.Explorer
    Set explorer = Application.ActiveExplorer
    If explorer Is Nothing Then Exit Function
    If explorer.Selection Is Nothing Then Exit Function
    If explorer.Selection.Count = 0 Then Exit Function

    If TypeOf explorer.Selection.Item(1) Is Outlook.MailItem Then
        Set GetCurrentMailItem = explorer.Selection.Item(1)
    End If
End Function

Private Function BuildConversationMailCollection(ByVal seed As Outlook.MailItem) As Collection
    On Error GoTo HandleConversationError

    Dim result As New Collection
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")

    Dim conv As Outlook.Conversation
    Set conv = seed.GetConversation

    If conv Is Nothing Then
        If MsgBox("У письма нет беседы (Conversation). Сделать summary только по этому письму?", vbYesNo + vbQuestion) = vbYes Then
            AddMailIfUnique seed, result, seen
            Set BuildConversationMailCollection = result
        End If
        Exit Function
    End If

    Dim roots As Outlook.SimpleItems
    Set roots = conv.GetRootItems

    Dim root As Object
    For Each root In roots
        WalkConversationNode conv, root, result, seen
    Next root

    If result.Count = 0 Then AddMailIfUnique seed, result, seen

    Set BuildConversationMailCollection = result
    Exit Function

HandleConversationError:
    AddMailIfUnique seed, result, seen
    Set BuildConversationMailCollection = result
End Function

Private Sub WalkConversationNode(ByVal conv As Outlook.Conversation, ByVal node As Object, ByRef output As Collection, ByRef seen As Object)
    On Error Resume Next

    If TypeOf node Is Outlook.MailItem Then
        AddMailIfUnique node, output, seen
    End If

    Dim children As Outlook.SimpleItems
    Set children = conv.GetChildren(node)
    If children Is Nothing Then Exit Sub

    Dim child As Object
    For Each child In children
        WalkConversationNode conv, child, output, seen
    Next child
End Sub

Private Sub AddMailIfUnique(ByVal item As Outlook.MailItem, ByRef output As Collection, ByRef seen As Object)
    On Error GoTo SafeExit
    If item Is Nothing Then Exit Sub

    Dim key As String
    key = item.EntryID & "|" & item.Parent.StoreID
    If Len(item.EntryID) = 0 Then key = "TEMP|" & item.Subject & "|" & CStr(item.ReceivedTime)

    If Not seen.Exists(key) Then
        seen.Add key, True
        output.Add item
    End If

SafeExit:
End Sub

Private Function SortMailsByDateAsc(ByVal mails As Collection) As Collection
    Dim arr() As Object
    ReDim arr(1 To mails.Count)

    Dim i As Long, j As Long
    For i = 1 To mails.Count
        Set arr(i) = mails(i)
    Next i

    For i = 1 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If MailDate(arr(j)) < MailDate(arr(i)) Then
                Dim tmp As Object
                Set tmp = arr(i)
                Set arr(i) = arr(j)
                Set arr(j) = tmp
            End If
        Next j
    Next i

    Dim sorted As New Collection
    For i = 1 To UBound(arr)
        sorted.Add arr(i)
    Next i

    Set SortMailsByDateAsc = sorted
End Function

Private Function MailDate(ByVal m As Outlook.MailItem) As Date
    On Error GoTo fallback
    If m.SentOn <> #1/1/4501# Then
        MailDate = m.SentOn
    Else
        MailDate = m.ReceivedTime
    End If
    Exit Function
fallback:
    MailDate = Now
End Function

Private Function BuildUserPrompt(ByVal mails As Collection) As String
    Dim sb As String
    Dim i As Long

    For i = 1 To mails.Count
        Dim m As Outlook.MailItem
        Set m = mails(i)

        Dim senderValue As String
        senderValue = m.SenderName
        If Len(Trim$(senderValue)) = 0 Then senderValue = m.SenderEmailAddress

        sb = sb & "--- Письмо " & CStr(i) & " ---" & vbCrLf
        sb = sb & "От: " & senderValue & vbCrLf
        sb = sb & "Кому: " & m.To & vbCrLf
        sb = sb & "Дата: " & Format$(MailDate(m), "yyyy-mm-dd HH:nn:ss") & vbCrLf
        sb = sb & "Тема: " & m.Subject & vbCrLf
        sb = sb & "Текст:" & vbCrLf
        sb = sb & CleanMailBody(m.Body) & vbCrLf & vbCrLf
    Next i

    BuildUserPrompt = sb
End Function

Private Function CleanMailBody(ByVal textBody As String) As String
    ' Историю переписки внутри письма НЕ отсекаем, чтобы модель могла
    ' дополнительно анализировать контекст, который пришел как цитирование.
    Dim s As String
    s = Replace(textBody, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)

    Dim lines() As String
    lines = Split(s, vbLf)

    Dim out As String
    Dim i As Long
    Dim previousWasEmpty As Boolean

    For i = LBound(lines) To UBound(lines)
        Dim line As String
        line = Trim$(lines(i))

        If Len(line) = 0 Then
            If Not previousWasEmpty Then
                out = out & vbCrLf
                previousWasEmpty = True
            End If
        Else
            out = out & line & vbCrLf
            previousWasEmpty = False
        End If
    Next i

    CleanMailBody = Trim$(out)
End Function

Private Function EnforceContextLimit(ByVal payload As String, ByRef warningMessage As String) As String
    warningMessage = ""

    If Len(payload) <= MAX_CONTEXT_CHARS Then
        EnforceContextLimit = payload
        Exit Function
    End If

    warningMessage = "Внимание: цепочка слишком длинная, отправлена только последняя часть переписки."
    EnforceContextLimit = Right$(payload, MAX_CONTEXT_CHARS)
End Function

Private Function SystemPrompt() As String
    SystemPrompt = "Ты — ассистент, который анализирует цепочки деловой переписки. " & _
                   "Перед тобой список писем в хронологическом порядке (от самого старого к самому новому). " & _
                   "Составь краткое саммари всей цепочки: выдели основных участников, опиши начальную проблему/вопрос, " & _
                   "ключевые аргументы и итоговое решение или текущее состояние дел на момент последнего письма. " & _
                   "Используй русский язык (или язык большинства писем). Будь краток и структурирован."
End Function

Private Function RequestSummaryFromLlm(ByVal userPayload As String) As String
    Dim apiKey As String
    apiKey = Trim$(Environ$(ENV_API_KEY))
    If Len(apiKey) = 0 Then Err.Raise vbObjectError + 2000, , "Не задан ключ API в переменной окружения " & ENV_API_KEY

    Dim apiUrl As String
    apiUrl = Trim$(Environ$(ENV_API_URL))
    If Len(apiUrl) = 0 Then apiUrl = DEFAULT_API_URL

    Dim modelName As String
    modelName = Trim$(Environ$(ENV_MODEL))
    If Len(modelName) = 0 Then modelName = DEFAULT_MODEL

    Dim json As String
    json = "{""model"":""" & JsonEscape(modelName) & """,""messages"": [" & _
           "{""role"":""system"",""content"":""" & JsonEscape(SystemPrompt()) & """}," & _
           "{""role"":""user"",""content"":""" & JsonEscape(userPayload) & """}]," & _
           ""temperature"":0.2}"

    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts REQUEST_TIMEOUT_MS, REQUEST_TIMEOUT_MS, REQUEST_TIMEOUT_MS, REQUEST_TIMEOUT_MS

    On Error GoTo NetworkError
    http.Open "POST", apiUrl, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & apiKey
    http.send json

    If http.Status < 200 Or http.Status >= 300 Then
        Err.Raise vbObjectError + 2001, , "Ошибка API: HTTP " & CStr(http.Status) & " - " & Left$(http.responseText, 600)
    End If

    RequestSummaryFromLlm = ExtractSummaryFromResponse(http.responseText)
    Exit Function

NetworkError:
    Err.Raise vbObjectError + 2002, , "Сетевая ошибка при запросе к LLM API: " & Err.Description
End Function

Private Function ExtractSummaryFromResponse(ByVal json As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.MultiLine = True
    re.Pattern = """content""\s*:\s*""((?:[^""\\]|\\.)*)"""

    Dim matches As Object
    Set matches = re.Execute(json)
    If matches.Count = 0 Then
        ExtractSummaryFromResponse = ""
        Exit Function
    End If

    Dim raw As String
    raw = matches(matches.Count - 1).SubMatches(0)
    ExtractSummaryFromResponse = JsonUnescape(raw)
End Function

Private Function JsonEscape(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", Chr$(92) & """)
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    JsonEscape = s
End Function

Private Function JsonUnescape(ByVal s As String) As String
    s = Replace(s, "\n", vbCrLf)
    s = Replace(s, "\r", vbCr)
    s = Replace(s, "\t", vbTab)
    s = Replace(s, Chr$(92) & """", """)
    s = Replace(s, "\\", "\")
    JsonUnescape = s
End Function

Private Function SaveSummaryToDesktop(ByVal summaryText As String) As String
    Dim shellObj As Object
    Set shellObj = CreateObject("WScript.Shell")

    Dim desktopPath As String
    desktopPath = shellObj.SpecialFolders("Desktop")

    Dim filePath As String
    filePath = desktopPath & "\" & Format$(Now, "yyyymmdd-Hhnn") & "-summary.txt"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim stream As Object
    Set stream = fso.CreateTextFile(filePath, True, True)
    stream.Write summaryText
    stream.Close

    SaveSummaryToDesktop = filePath
End Function
