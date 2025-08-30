Attribute VB_Name = "WordiCalc"
Option Explicit

' WordiCalc - AI-Powered Excel Extension
' Adds =LLM(...) function for OpenAI integration
' Version: 1.1

Private Const DEFAULT_API_URL As String = "https://api.openai.com/v1/chat/completions"
Private Const CONFIG_PREFIX As String = "LLM_"

' ========================================
' MAIN FUNCTIONS
' ========================================

Public Function LLM(prompt As String, Optional sys As String = "You are a helpful assistant.", Optional schema As String = "string", Optional values As String = "", Optional useJson As Boolean = False) As Variant
    Application.Volatile
    On Error GoTo ErrorHandler
    
    If Trim(prompt) = "" Then LLM = "Error: Prompt cannot be empty": Exit Function
    
    schema = LCase(Trim(schema))
    If schema <> "string" And schema <> "integer" And schema <> "float" And schema <> "choice" Then
        LLM = "Error: Invalid schema. Use: string, integer, float, choice": Exit Function
    End If
    If schema = "choice" And values = "" Then LLM = "Error: Choice schema requires allowed values": Exit Function
    
    ' Enhance system message for non-JSON mode
    If Not useJson Then
        Select Case schema
            Case "integer": sys = sys & " Respond with only a single integer."
            Case "float": sys = sys & " Respond with only a single number."
            Case "choice": sys = sys & " Respond with exactly one of: " & values & ". No other text."
        End Select
    End If
    
    Dim response As String
    response = CallAPI(prompt, sys, schema, values, useJson)
    
    If Left(response, 6) = "Error:" Then LLM = response: Exit Function
    LLM = ConvertOutput(response, schema, values, useJson)
    Exit Function
    
ErrorHandler:
    LLM = "Error: " & Err.Description
End Function

Public Function LLMConfig(action As String, Optional key As String = "", Optional value As String = "") As String
    Application.Volatile
    On Error GoTo ErrorHandler
    
    Select Case LCase(Trim(action))
        Case "set"
            If key = "" Or value = "" Then LLMConfig = "Error: Key and value required": Exit Function
            SetConfig key, value: LLMConfig = "Configuration set"
        Case "get": LLMConfig = GetConfig(key)
        Case "list": LLMConfig = ListConfigs()
        Case "clear": ClearConfig key: LLMConfig = "Configuration cleared"
        Case Else: LLMConfig = "Error: Use set, get, list, or clear"
    End Select
    Exit Function
    
ErrorHandler:
    LLMConfig = "Error: " & Err.Description
End Function

Public Function LLMStatus() As String
    Application.Volatile
    Dim apiKey As String, endpoint As String, model As String
    apiKey = GetConfig("openai_api_key")
    endpoint = GetConfig("openai_api_endpoint")
    model = GetConfig("openai_model")
    
    LLMStatus = "WordiCalc v1.1" & vbCrLf & _
                "API Key: " & IIf(apiKey <> "", "Configured", "Not configured") & vbCrLf & _
                "Endpoint: " & IIf(endpoint <> "", endpoint, DEFAULT_API_URL & " (default)") & vbCrLf & _
                "Model: " & IIf(model <> "", model, "gpt-3.5-turbo (default)")
End Function

Public Function LLMModels() As String
    Application.Volatile
    On Error GoTo ErrorHandler
    
    If GetConfig("openai_api_key") = "" Then
        LLMModels = "Error: API key not configured"
        Exit Function
    End If
    
    Dim endpoint As String
    endpoint = GetConfig("openai_api_endpoint")
    If endpoint = "" Then endpoint = DEFAULT_API_URL
    
    ' Convert to models endpoint
    endpoint = Replace(endpoint, "/chat/completions", "/models")
    If Right(endpoint, 7) <> "/models" Then endpoint = endpoint & "/models"
    
    Dim headers As Object
    Set headers = CreateObject("Scripting.Dictionary")
    headers("Authorization") = "Bearer " & GetConfig("openai_api_key")
    
    Dim result As Object
    Set result = HttpRequest(endpoint, "GET", headers, "")
    
    If Not result("success") Then
        LLMModels = "Error: " & result("status") & " " & result("statusText")
        Exit Function
    End If
    
    LLMModels = ParseModels(result("responseText"))
    Exit Function
    
ErrorHandler:
    LLMModels = "Error: " & Err.Description
End Function

' ========================================
' CORE FUNCTIONS
' ========================================

Private Function CallAPI(prompt As String, sys As String, schema As String, values As String, useJson As Boolean) As String
    On Error GoTo ErrorHandler
    
    If GetConfig("openai_api_key") = "" Then
        CallAPI = "Error: API key not configured. Use =LLMConfig(""set"", ""openai_api_key"", ""your-key"")"
        Exit Function
    End If
    
    Dim model As String, endpoint As String
    model = GetConfig("openai_model"): If model = "" Then model = "gpt-3.5-turbo"
    endpoint = GetConfig("openai_api_endpoint"): If endpoint = "" Then endpoint = DEFAULT_API_URL
    
    Dim headers As Object
    Set headers = CreateObject("Scripting.Dictionary")
    headers("Content-Type") = "application/json"
    headers("Authorization") = "Bearer " & GetConfig("openai_api_key")
    
    Dim messages As String
    messages = "["
    If sys <> "" Then messages = messages & "{""role"":""system"",""content"":""" & EscapeJson(sys) & """},"
    messages = messages & "{""role"":""user"",""content"":""" & EscapeJson(prompt) & """}]"
    
    Dim body As String
    body = "{""model"":""" & model & """,""max_tokens"":1000,""messages"":" & messages
    
    ' Add JSON schema if needed
    If useJson And schema <> "string" Then
        Dim responseFormat As String
        responseFormat = BuildSchema(schema, values)
        If responseFormat <> "" Then body = body & ",""response_format"":" & responseFormat
    End If
    
    body = body & "}"
    
    Dim result As Object
    Set result = HttpRequest(endpoint, "POST", headers, body)
    
    If Not result("success") Then
        CallAPI = "Error: " & result("status") & " " & result("statusText")
        Exit Function
    End If
    
    CallAPI = ExtractContent(result("responseText"), useJson, schema)
    Exit Function
    
ErrorHandler:
    CallAPI = "Error: " & Err.Description
End Function

Private Function ConvertOutput(response As String, schema As String, values As String, useJson As Boolean) As Variant
    On Error GoTo ErrorHandler
    
    response = Trim(response)
    
    Select Case schema
        Case "string": ConvertOutput = response
        Case "integer", "float"
            Dim numStr As String
            If useJson Or IsNumeric(response) Then
                numStr = response
            Else
                numStr = ExtractNumber(response)
            End If
            If IsNumeric(numStr) Then
                ConvertOutput = Application.WorksheetFunction.NumberValue(numStr)
            Else
                ConvertOutput = "Error: No valid number found in: " & response
            End If
        Case "choice"
            Dim choices As Variant: choices = Split(values, ",")
            Dim i As Integer
            For i = 0 To UBound(choices)
                Dim choice As String: choice = Trim(choices(i))
                If InStr(LCase(response), LCase(choice)) > 0 Then
                    ConvertOutput = choice: Exit Function
                End If
            Next i
            ConvertOutput = "Error: Response '" & response & "' not in: " & values
        Case Else: ConvertOutput = response
    End Select
    Exit Function
    
ErrorHandler:
    ConvertOutput = "Error: " & Err.Description
End Function

' ========================================
' UTILITY FUNCTIONS
' ========================================

Private Function HttpRequest(url As String, method As String, headers As Object, body As String) As Object
    Dim http As Object: Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    Dim result As Object: Set result = CreateObject("Scripting.Dictionary")
    
    On Error GoTo ErrorHandler
    
    http.Open method, url, False
    http.SetTimeouts 30000, 30000, 30000, 30000
    
    If Not headers Is Nothing Then
        Dim key As Variant
        For Each key In headers.Keys
            http.SetRequestHeader CStr(key), CStr(headers(key))
        Next key
    End If
    
    http.Send body
    
    result("status") = http.Status
    result("statusText") = http.StatusText
    result("responseText") = http.ResponseText
    result("success") = (http.Status >= 200 And http.Status < 300)
    
    Set HttpRequest = result
    Exit Function
    
ErrorHandler:
    result("status") = 0
    result("statusText") = "Request Failed"
    result("responseText") = "Error: " & Err.Description
    result("success") = False
    Set HttpRequest = result
End Function

Private Function ExtractContent(jsonResponse As String, useJson As Boolean, schema As String) As String
    On Error GoTo ErrorHandler
    
    Dim parsed As Object: Set parsed = ParseJson(jsonResponse)
    If parsed Is Nothing Then ExtractContent = "Error: Invalid JSON response": Exit Function
    
    If parsed.Exists("error") Then
        ExtractContent = "API Error: " & GetJsonValue(parsed("error"), "message")
        Exit Function
    End If
    
    If Not parsed.Exists("choices") Then
        ExtractContent = "Error: No choices in response": Exit Function
    End If
    
    Dim content As String: content = ExtractFirstChoice(parsed("choices"))
    If content = "" Then ExtractContent = "Error: No content found": Exit Function
    
    ' Handle JSON schema response
    If useJson And schema <> "string" Then
        Dim contentJson As Object: Set contentJson = ParseJson(content)
        If Not contentJson Is Nothing And contentJson.Exists("value") Then
            content = CStr(contentJson("value"))
        End If
    End If
    
    ExtractContent = content
    Exit Function
    
ErrorHandler:
    ExtractContent = "Error: " & Err.Description
End Function

Private Function ExtractFirstChoice(choicesStr As String) As String
    Dim pos As Long: pos = InStr(choicesStr, """content"":""")
    If pos = 0 Then Exit Function
    
    pos = pos + 11 ' Skip to after "content":"
    Dim content As String, i As Long, inEscape As Boolean
    
    For i = pos To Len(choicesStr)
        Dim c As String: c = Mid(choicesStr, i, 1)
        If inEscape Then
            content = content & c: inEscape = False
        ElseIf c = "\" Then
            inEscape = True
        ElseIf c = """" Then
            Exit For
        Else
            content = content & c
        End If
    Next i
    
    ExtractFirstChoice = Replace(Replace(content, "\""", """"), "\\", "\")
End Function

Private Function ParseJson(jsonStr As String) As Object
    ' Simple JSON parser for basic objects
    Set ParseJson = CreateObject("Scripting.Dictionary")
    On Error GoTo ErrorHandler
    
    jsonStr = Trim(jsonStr)
    If Left(jsonStr, 1) <> "{" Then Set ParseJson = Nothing: Exit Function
    
    Dim i As Long, key As String, value As String, inString As Boolean, depth As Long
    Dim parsingKey As Boolean: parsingKey = True
    i = 2
    
    Do While i <= Len(jsonStr) - 1
        Dim c As String: c = Mid(jsonStr, i, 1)
        
        If c = """" Then
            inString = Not inString
        ElseIf Not inString Then
            If c = "{" Or c = "[" Then depth = depth + 1
            If c = "}" Or c = "]" Then depth = depth - 1
            If depth = 0 And c = ":" Then parsingKey = False
            If depth = 0 And c = "," Then
                If key <> "" Then
                    key = Replace(key, """", "")
                    value = Trim(value)
                    If Left(value, 1) = """" And Right(value, 1) = """" Then
                        value = Mid(value, 2, Len(value) - 2)
                    End If
                    ParseJson(key) = value
                End If
                key = "": value = "": parsingKey = True
            Else
                If parsingKey And c <> " " And c <> vbTab Then key = key & c
                If Not parsingKey Then value = value & c
            End If
        Else
            If parsingKey Then key = key & c Else value = value & c
        End If
        i = i + 1
    Loop
    
    ' Handle last key-value pair
    If key <> "" Then
        key = Replace(key, """", "")
        value = Trim(value)
        If Left(value, 1) = """" And Right(value, 1) = """" Then
            value = Mid(value, 2, Len(value) - 2)
        End If
        ParseJson(key) = value
    End If
    Exit Function
    
ErrorHandler:
    Set ParseJson = Nothing
End Function

Private Function GetJsonValue(jsonObj As String, keyName As String) As String
    Dim obj As Object: Set obj = ParseJson(jsonObj)
    If Not obj Is Nothing And obj.Exists(keyName) Then
        GetJsonValue = obj(keyName)
    Else
        GetJsonValue = jsonObj ' Return as-is if can't parse
    End If
End Function

Private Function BuildSchema(schema As String, values As String) As String
    On Error GoTo ErrorHandler
    
    Dim schemaObj As String
    Select Case schema
        Case "integer": schemaObj = "{""type"":""object"",""properties"":{""value"":{""type"":""integer""}},""required"":[""value""]}"
        Case "float": schemaObj = "{""type"":""object"",""properties"":{""value"":{""type"":""number""}},""required"":[""value""]}"
        Case "choice"
            If values = "" Then Exit Function
            Dim choices As Variant: choices = Split(values, ",")
            Dim enumStr As String: enumStr = "["
            Dim i As Integer
            For i = 0 To UBound(choices)
                If i > 0 Then enumStr = enumStr & ","
                enumStr = enumStr & """" & Trim(choices(i)) & """"
            Next i
            enumStr = enumStr & "]"
            schemaObj = "{""type"":""object"",""properties"":{""value"":{""type"":""string"",""enum"":" & enumStr & "}},""required"":[""value""]}"
        Case Else: Exit Function
    End Select
    
    BuildSchema = "{""type"":""json_schema"",""json_schema"":{""name"":""response"",""schema"":" & schemaObj & "}}"
    Exit Function
    
ErrorHandler:
    BuildSchema = ""
End Function

Private Function ParseModels(jsonResponse As String) As String
    On Error GoTo ErrorHandler
    
    Dim parsed As Object: Set parsed = ParseJson(jsonResponse)
    If parsed Is Nothing Then ParseModels = "Error: Could not parse response": Exit Function
    
    Dim result As String: result = "Available Models:" & vbCrLf
    Dim modelCount As Integer
    
    If parsed.Exists("data") Then
        Dim modelsStr As String: modelsStr = parsed("data")
        Dim pos As Long: pos = 1
        Do While pos < Len(modelsStr)
            pos = InStr(pos, modelsStr, """id"":")
            If pos = 0 Then Exit Do
            pos = InStr(pos, modelsStr, """") + 1
            Dim modelEnd As Long: modelEnd = InStr(pos, modelsStr, """")
            If modelEnd > pos Then
                Dim modelName As String: modelName = Mid(modelsStr, pos, modelEnd - pos)
                If InStr(modelName, "embed") = 0 And InStr(modelName, "whisper") = 0 Then
                    result = result & "• " & modelName & vbCrLf
                    modelCount = modelCount + 1
                End If
            End If
            pos = modelEnd + 1
        Loop
    End If
    
    If modelCount = 0 Then
        result = result & "No models found. Raw response: " & Left(jsonResponse, 200)
    End If
    
    ParseModels = result
    Exit Function
    
ErrorHandler:
    ParseModels = "Error: " & Err.Description
End Function

Private Function ExtractNumber(text As String) As String
    Dim result As String, foundDecimal As Boolean
    Dim i As Integer
    For i = 1 To Len(text)
        Dim c As String: c = Mid(text, i, 1)
        If IsNumeric(c) Then
            result = result & c
        ElseIf c = "." And Not foundDecimal And result <> "" Then
            result = result & c: foundDecimal = True
        ElseIf c = "-" And result = "" Then
            result = result & c
        ElseIf result <> "" Then
            Exit For
        End If
    Next i
    ExtractNumber = result
End Function

Private Function EscapeJson(text As String) As String
    EscapeJson = text
    EscapeJson = Replace(EscapeJson, "\", "\\")
    EscapeJson = Replace(EscapeJson, """", "\""")
    EscapeJson = Replace(EscapeJson, vbCr, "\r")
    EscapeJson = Replace(EscapeJson, vbLf, "\n")
    EscapeJson = Replace(EscapeJson, vbTab, "\t")
End Function

' ========================================
' CONFIG FUNCTIONS
' ========================================

Private Sub SetConfig(key As String, value As String)
    On Error GoTo ErrorHandler
    key = CONFIG_PREFIX & key
    
    Dim prop As DocumentProperty, found As Boolean
    For Each prop In ThisWorkbook.CustomDocumentProperties
        If prop.Name = key Then prop.value = value: found = True: Exit For
    Next prop
    
    If Not found Then
        ThisWorkbook.CustomDocumentProperties.Add Name:=key, LinkToContent:=False, Type:=msoPropertyTypeString, value:=value
    End If
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, "SetConfig", "Failed: " & Err.Description
End Sub

Private Function GetConfig(key As String) As String
    On Error GoTo ErrorHandler
    key = CONFIG_PREFIX & key
    
    Dim prop As DocumentProperty
    For Each prop In ThisWorkbook.CustomDocumentProperties
        If prop.Name = key Then GetConfig = prop.value: Exit Function
    Next prop
    
    GetConfig = ""
    Exit Function
    
ErrorHandler:
    GetConfig = ""
End Function

Private Sub ClearConfig(key As String)
    On Error GoTo ErrorHandler
    key = CONFIG_PREFIX & key
    
    Dim prop As DocumentProperty
    For Each prop In ThisWorkbook.CustomDocumentProperties
        If prop.Name = key Then prop.Delete: Exit For
    Next prop
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, "ClearConfig", "Failed: " & Err.Description
End Sub

Private Function ListConfigs() As String
    On Error GoTo ErrorHandler
    
    Dim result As String: result = "Configuration:" & vbCrLf
    Dim found As Boolean
    
    Dim prop As DocumentProperty
    For Each prop In ThisWorkbook.CustomDocumentProperties
        If Left(prop.Name, Len(CONFIG_PREFIX)) = CONFIG_PREFIX Then
            Dim displayKey As String: displayKey = Mid(prop.Name, Len(CONFIG_PREFIX) + 1)
            Dim displayValue As String
            If InStr(LCase(displayKey), "key") > 0 Then
                displayValue = "****** (hidden)"
            Else
                displayValue = prop.value
            End If
            result = result & displayKey & ": " & displayValue & vbCrLf
            found = True
        End If
    Next prop
    
    If Not found Then result = result & "No configurations found."
    ListConfigs = result
    Exit Function
    
ErrorHandler:
    ListConfigs = "Error: " & Err.Description
End Function