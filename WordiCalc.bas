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
    On Error GoTo ErrorHandler
    
    Dim validationError As String
    validationError = ValidateParameters(prompt, schema, values)
    If validationError <> "" Then LLM = validationError: Exit Function
    schema = LCase(Trim(schema))
    
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
    
    Dim headers As Dictionary
    Set headers = New Dictionary
    headers("Authorization") = "Bearer " & GetConfig("openai_api_key")
    
    Dim result As Dictionary
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
    
    Dim headers As Dictionary
    Set headers = New Dictionary
    headers("Content-Type") = "application/json"
    headers("Authorization") = "Bearer " & GetConfig("openai_api_key")
    
    Dim body As String
    body = BuildRequestBody(model, prompt, sys, schema, values, useJson)
    
    Dim result As Dictionary
    Set result = HttpRequest(endpoint, "POST", headers, body)
    
    If Not result("success") Then
        CallAPI = "Error: " & result("status") & " " & result("statusText") & vbCrLf & _
                 "DEBUG INFO: " & result("debug_info")
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
            ConvertOutput = ConvertToNumber(response, schema, useJson)
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

Private Function ConvertToNumber(response As String, schema As String, useJson As Boolean) As Variant
    Dim numStr As String
    If useJson Or IsNumeric(response) Then
        numStr = response
    Else
        numStr = ExtractNumber(response)
    End If
    
    If IsNumeric(numStr) Then
        If schema = "integer" Then
            ConvertToNumber = CLng(Val(numStr))
        Else
            ConvertToNumber = CDbl(Val(numStr))
        End If
    Else
        ConvertToNumber = "Error: No valid number found in: " & response
    End If
End Function

Private Function ValidateParameters(prompt As String, schema As String, values As String) As String
    If Trim(prompt) = "" Then
        ValidateParameters = "Error: Prompt cannot be empty"
        Exit Function
    End If
    
    Dim schemaLower As String: schemaLower = LCase(Trim(schema))
    If schemaLower <> "string" And schemaLower <> "integer" And schemaLower <> "float" And schemaLower <> "choice" Then
        ValidateParameters = "Error: Invalid schema. Use: string, integer, float, choice"
        Exit Function
    End If
    
    If schemaLower = "choice" And values = "" Then
        ValidateParameters = "Error: Choice schema requires allowed values"
        Exit Function
    End If
    
    ValidateParameters = "" ' No errors
End Function

' ========================================
' UTILITY FUNCTIONS
' ========================================

Private Function HttpRequest(url As String, method As String, headers As Dictionary, body As String) As Dictionary
    Dim http As Object: Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    Dim result As Dictionary: Set result = New Dictionary
    
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
    
    ' DEBUG: Log raw response for debugging
    result("debug_info") = "URL: " & url & vbCrLf & _
                          "Status: " & http.Status & vbCrLf & _
                          "Response Length: " & Len(http.ResponseText) & vbCrLf & _
                          "First 200 chars: " & Left(http.ResponseText, 200)
    
    Set HttpRequest = result
    Exit Function
    
ErrorHandler:
    result("status") = 0
    result("statusText") = "Request Failed"
    result("responseText") = "Error: " & Err.Description
    result("success") = False
    result("debug_info") = "HTTP Error: " & Err.Description
    Set HttpRequest = result
End Function

Private Function ExtractContent(jsonResponse As String, useJson As Boolean, schema As String) As String
    On Error GoTo ErrorHandler
    
    Dim parseResult As ParseResult: parseResult = LibJSON.Parse(jsonResponse)
    If Not parseResult.IsValid Then
        ExtractContent = "JSON Parse Error: " & parseResult.Error
        Exit Function
    End If
    Dim parsed As Object: Set parsed = parseResult.Value
    
    If parsed.Exists("error") Then
        If TypeName(parsed("error")) = "Dictionary" Then
            ExtractContent = "API Error: " & parsed("error")("message")
        Else
            ExtractContent = "API Error: " & CStr(parsed("error"))
        End If
        Exit Function
    End If
    
    If Not parsed.Exists("choices") Then
        ExtractContent = "Error: No choices in response"
        Exit Function
    End If
    
    ' Get the first choice from the choices array/collection
    Dim choices As Object: Set choices = parsed("choices")
    Dim firstChoice As Object
    
    If TypeName(choices) = "Collection" Then
        Set firstChoice = choices(1)  ' Collections are 1-based
    Else
        ' Assume it's an array or dictionary
        Set firstChoice = choices(0)  ' Arrays are 0-based
    End If
    
    Dim content As String
    If firstChoice.Exists("message") Then
        content = firstChoice("message")("content")
    Else
        content = firstChoice("content")
    End If
    
    If content = "" Then
        ExtractContent = "Error: No content found in first choice"
        Exit Function
    End If
    
    ' Handle JSON schema response
    If useJson And schema <> "string" Then
        Dim contentParseResult As ParseResult: contentParseResult = LibJSON.Parse(content)
        If Not contentParseResult.IsValid Then Exit Function
        Dim contentJson As Object: Set contentJson = contentParseResult.Value
        If Not contentJson Is Nothing And contentJson.Exists("value") Then
            content = CStr(contentJson("value"))
        End If
    End If
    
    ExtractContent = content
    Exit Function
    
ErrorHandler:
    ExtractContent = "Error: " & Err.Description
End Function


Private Function BuildRequestBody(model As String, prompt As String, sys As String, schema As String, values As String, useJson As Boolean) As String
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
    
    BuildRequestBody = body & "}"
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
    
    Dim parseResult As ParseResult: parseResult = LibJSON.Parse(jsonResponse)
    If Not parseResult.IsValid Then
        ParseModels = "Error: Could not parse response - " & parseResult.Error
        Exit Function
    End If
    Dim parsed As Object: Set parsed = parseResult.Value
    
    Dim result As String: result = "Available Models:" & vbCrLf
    Dim modelCount As Integer
    
    If parsed.Exists("data") Then
        Dim models As Object: Set models = parsed("data")
        Dim i As Long
        For i = 1 To models.Count ' LibJSON collections are 1-based
            Dim model As Object: Set model = models(i)
            If model.Exists("id") Then
                Dim modelName As String: modelName = model("id")
                If InStr(modelName, "embed") = 0 And InStr(modelName, "whisper") = 0 Then
                    result = result & "• " & modelName & vbCrLf
                    modelCount = modelCount + 1
                End If
            End If
        Next i
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
' FUNCTION REGISTRATION
' ========================================

Public Sub RegisterUDFFunctions()
    ' Register LLM function with IntelliSense support
    On Error Resume Next
    
    Dim llmArgDesc(1 To 5) As String
    llmArgDesc(1) = "The text prompt to send to the AI model"
    llmArgDesc(2) = "System message to guide AI behavior (optional, default: 'You are a helpful assistant.')"
    llmArgDesc(3) = "Response format: string, integer, float, choice (optional, default: 'string')"
    llmArgDesc(4) = "Allowed values for choice schema, comma-separated (optional, required for choice schema)"
    llmArgDesc(5) = "Use JSON schema mode for structured output (optional, default: False)"
    
    Application.MacroOptions _
        Macro:="LLM", _
        Description:="Call OpenAI/Compatible API with prompt and optional parameters for AI-powered responses", _
        Category:="WordiCalc AI Functions", _
        ArgumentDescriptions:=llmArgDesc
    
    ' Register LLMConfig function
    Dim configArgDesc(1 To 3) As String
    configArgDesc(1) = "Action to perform: set, get, list, or clear"
    configArgDesc(2) = "Configuration key name (optional, required for set/get/clear)"
    configArgDesc(3) = "Configuration value (optional, required for set action)"
    
    Application.MacroOptions _
        Macro:="LLMConfig", _
        Description:="Manage WordiCalc configuration settings (API key, endpoint, model)", _
        Category:="WordiCalc AI Functions", _
        ArgumentDescriptions:=configArgDesc
    
    ' Register LLMStatus function
    Application.MacroOptions _
        Macro:="LLMStatus", _
        Description:="Display current WordiCalc configuration status and version information", _
        Category:="WordiCalc AI Functions"
    
    ' Register LLMModels function
    Application.MacroOptions _
        Macro:="LLMModels", _
        Description:="List available AI models from the configured API endpoint", _
        Category:="WordiCalc AI Functions"
    
    On Error GoTo 0
End Sub

Public Sub Auto_Open()
    ' Automatically register UDF functions when workbook opens
    RegisterUDFFunctions
End Sub


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