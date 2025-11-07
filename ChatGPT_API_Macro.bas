Sub ChatGPT_CustomPrompt_InsertBelow_Fixed()
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    Dim url As String
    url = "https://api.openai.com/v1/chat/completions"
    Dim apiKey As String
    apiKey = "YOUR API KEY"
    
    If Selection.Text = "" Then
        MsgBox "Bitte zuerst Text markieren."
        Exit Sub
    End If
    
    Dim userPrompt As String
    userPrompt = InputBox("Was soll ChatGPT mit dem markierten Text machen?", "ChatGPT Prompt", "Fasse diesen Text kurz zusammen:")
    If userPrompt = "" Then Exit Sub
    
    Dim prompt As String
    prompt = userPrompt & " " & Selection.Text
    
    Dim json As String
    json = "{""model"":""gpt-4o-mini"",""messages"":[{""role"":""user"",""content"":""" & JsonEscape(prompt) & """}]}"
    
    With http
        .Open "POST", url, False
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Authorization", "Bearer " & apiKey
        .Send json
    End With
    
    Dim response As String
    response = http.responseText
    
    Dim result As String
    result = ExtractAssistantContent(response)
    
    If result <> "" Then
        result = Replace(result, "\n", vbCrLf)
        result = Replace(result, "\\", "\")
        result = Replace(result, "\""", """")
        Selection.Collapse Direction:=wdCollapseEnd
        Selection.TypeParagraph
        Selection.TypeText Text:="?? ChatGPT (" & userPrompt & "):" & vbCrLf & result
    Else
        MsgBox "Fehler beim Lesen der Antwort:" & vbCrLf & vbCrLf & response
    End If
End Sub

Function JsonEscape(s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    JsonEscape = s
End Function

Function ExtractAssistantContent(json As String) As String
    Dim contentKey As String
    contentKey = """content"":"
    
    Dim pos As Long
    pos = InStr(json, contentKey)
    
    If pos > 0 Then
        Dim startPos As Long
        startPos = pos + Len(contentKey)
        
        ' Finde das erste Anführungszeichen nach "content":
        Do While Mid(json, startPos, 1) <> """"
            startPos = startPos + 1
            If startPos > Len(json) Then Exit Function
        Loop
        startPos = startPos + 1
        
        ' Suche das nächste nicht-escaped Anführungszeichen:
        Dim i As Long
        For i = startPos To Len(json)
            If Mid(json, i, 1) = """" And Mid(json, i - 1, 1) <> "\" Then
                ExtractAssistantContent = Mid(json, startPos, i - startPos)
                Exit Function
            End If
        Next i
    End If
End Function

