Attribute VB_Name = "Module1"
Option Explicit

Public Flg As LongPtr
Public Declare PtrSafe Function InternetGetConnectedState Lib "wininet.dll" (lpdwFlags As LongPtr, ByVal dwReserved As Long) As Boolean

Private Const INTERNET_CONNECTION_MODEM As Long = &H1
Private Const INTERNET_CONNECTION_LAN As Long = &H2
Private Const INTERNET_CONNECTION_PROXY As Long = &H4
Private Const INTERNET_CONNECTION_MODEM_BUSY As Long = &H8
Private Const INTERNET_RAS_INSTALLED As Long = &H10
Private Const INTERNET_CONNECTION_OFFLINE As Long = &H20
Private Const INTERNET_CONNECTION_CONFIGURED As Long = &H40
'-----------------------------------------------------------------------------------------------
Public Function IsInternetConnected() As Boolean
    Dim R As Long

    R = InternetGetConnectedState(Flg, 0&)

    If Flg >= INTERNET_CONNECTION_OFFLINE Then
        Debug.Print "INTERNET_CONNECTION_OFFLINE"
    End If

    If CBool(R) Then
        IsInternetConnected = True
    Else
        IsInternetConnected = False
    End If
    
End Function
'-----------------------------------------------------------------------------------------------
Public Sub Test_URL_ChatGPT()
    
    Dim strOpenAI_URL As String
    Dim strOpenAI_APIKey As String
    Dim strOpenAI_Model As String
    Dim bRes As Boolean
    
    ThisWorkbook.Sheets("VBA").TextBox9 = ""
    ThisWorkbook.Sheets("VBA").TextBox10 = ""
    
    strOpenAI_URL = ThisWorkbook.Sheets("VBA").TextBox8
    strOpenAI_APIKey = ThisWorkbook.Sheets("VBA").TextBox1
    strOpenAI_Model = ThisWorkbook.Sheets("VBA").ComboBox1
    
    bRes = IsInternetConnected()
    If (bRes = False) Then
        ThisWorkbook.Sheets("VBA").TextBox9 = "Internet: NOK"
        Exit Sub
    Else
        ThisWorkbook.Sheets("VBA").TextBox9 = "Internet: OK"
    End If
    
    
    bRes = OpenAI_Test_URL(strOpenAI_URL, strOpenAI_APIKey, strOpenAI_Model)
    
End Sub
'-----------------------------------------------------------------------------------------------
Public Sub Processing_ChatGPT()

    Dim strPrompt As String
    Dim strOpenAI_URL As String
    Dim strOpenAI_APIKey As String
    Dim strOpenAI_Model As String
    Dim bRes As Boolean
    Dim strStatus As String
    
    Application.ScreenUpdating = True
    ThisWorkbook.Sheets("VBA").TextBox5 = ""
    ThisWorkbook.Sheets("VBA").TextBox3 = ""
    ThisWorkbook.Sheets("VBA").TextBox4 = ""
    
    ThisWorkbook.Sheets("VBA").TextBox15 = ""
    ThisWorkbook.Sheets("VBA").TextBox16 = ""
    
    strOpenAI_URL = ThisWorkbook.Sheets("VBA").TextBox8
    strOpenAI_APIKey = ThisWorkbook.Sheets("VBA").TextBox1
    strOpenAI_Model = ThisWorkbook.Sheets("VBA").ComboBox1
    strPrompt = ThisWorkbook.Sheets("VBA").TextBox2
    
    bRes = IsInternetConnected()
    If (bRes = False) Then
        ThisWorkbook.Sheets("VBA").TextBox9 = "Internet: NOK"
        Exit Sub
    Else
        ThisWorkbook.Sheets("VBA").TextBox9 = "Internet: OK"
    End If
    
    
    If (strOpenAI_Model = "gpt-3.5-turbo-instruct" Or strOpenAI_Model = "babbage-002" Or strOpenAI_Model = "davinci-002") Then
        bRes = OpenAI_Completion(strOpenAI_URL, strOpenAI_APIKey, strOpenAI_Model, strPrompt)
        bRes = False
    ElseIf (Left(strOpenAI_Model, 4) = "gpt-") Then
        bRes = OpenAI_ChatCompletion(strOpenAI_URL, strOpenAI_APIKey, strOpenAI_Model, strPrompt)
    Else
        MsgBox ("The VBA code doesn't support this model yet.")
        bRes = False
    End If
    
    If bRes = False Then
        strStatus = "Error"
    Else
        strStatus = "Done"
    End If
        
    ThisWorkbook.Sheets("VBA").TextBox5 = strStatus
    
End Sub
'-----------------------------------------------------------------------------------------------
Private Function OpenAI_Completion(ByVal strOpenAI_URL As String, ByVal strOpenAI_APIKey As String, ByVal strOpenAI_Model As String, ByVal strInputTextUser As String) As Boolean

    Dim httpRequest As MSXML2.XMLHTTP60
    Dim strAuthCredentials As String
    Dim requestReadyState As Long
    Dim requestStatus As Long
    Dim requestJSON As String
    Dim strURL As String
    Dim completion As String
    Dim RESPONSE As String
    Dim tokens As String
    Dim pricing_per_token As Single
    Dim pricing As Single

    ThisWorkbook.Sheets("VBA").TextBox5 = "Processing"
    
    strURL = strOpenAI_URL & "/v1/completions"
    
    pricing_per_token = OpenAI_Pricing(strOpenAI_Model)
    
    strAuthCredentials = "Bearer " & strOpenAI_APIKey
    requestJSON = OpenAI_Model2JSON(strInputTextUser, strOpenAI_Model)
    
    
    On Error GoTo OpenAI_Completion
    
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    
    'Open and send the HTTP request
    Call httpRequest.Open("POST", strURL, False)
    Call httpRequest.SetRequestHeader("Content-Type", "application/json")
    Call httpRequest.SetRequestHeader("Authorization", strAuthCredentials)
    Call httpRequest.Send(requestJSON)
    
    requestReadyState = httpRequest.readyState
    
    
    'Check if the request is successful
    If requestReadyState = 1 Then
        ThisWorkbook.Sheets("VBA").TextBox4 = "Error: httpRequest.readyState=1"
        OpenAI_Completion = False
    ElseIf requestReadyState = 4 Then
        requestStatus = httpRequest.Status
        completion = httpRequest.responseText
        ThisWorkbook.Sheets("VBA").TextBox4 = completion
        ThisWorkbook.Sheets("VBA").TextBox11 = requestStatus
        If requestStatus = 200 Then
            ThisWorkbook.Sheets("VBA").TextBox10 = "General checks: OK"
            
            'Parse the JSON response
            RESPONSE = OpenAI_Decode_JSON_completion(completion, strOpenAI_Model)
            tokens = OpenAI_Decode_JSON_completion_tokens(completion)
            pricing = CInt(tokens) * pricing_per_token
            ThisWorkbook.Sheets("VBA").TextBox3 = RESPONSE
            ThisWorkbook.Sheets("VBA").TextBox15 = tokens & " tokens"
            ThisWorkbook.Sheets("VBA").TextBox16 = "$ " & CStr(pricing)
            OpenAI_Completion = True
        Else
            'Parse the JSON response
            RESPONSE = OpenAI_Decode_JSON_ChatGPTError(completion)
            ThisWorkbook.Sheets("VBA").TextBox3 = RESPONSE
            OpenAI_Completion = False
        End If
    Else
        ThisWorkbook.Sheets("VBA").TextBox4 = "Error: httpRequest.readyState not defined"
        OpenAI_Completion = False
    End If
    
    
    
   On Error GoTo 0
    Exit Function
        
OpenAI_Completion:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure OpenAI_Completion, line " & Erl & "."

End Function
'-----------------------------------------------------------------------------------------------
Private Function OpenAI_ChatCompletion(ByVal strOpenAI_URL As String, ByVal strOpenAI_APIKey As String, ByVal strOpenAI_Model As String, ByVal strInputTextUser As String) As Boolean

    Dim httpRequest As MSXML2.XMLHTTP60
    Dim strAuthCredentials As String
    Dim requestReadyState As Long
    Dim requestStatus As Long
    Dim requestJSON As String
    Dim strURL As String
    Dim completion As String
    Dim RESPONSE As String
    Dim tokens As String
    Dim pricing_per_token As Single
    Dim pricing As Single
    
    ThisWorkbook.Sheets("VBA").TextBox5 = "Processing"
    
    strURL = strOpenAI_URL & "/v1/chat/completions"
    
    pricing_per_token = OpenAI_Pricing(strOpenAI_Model)
    
    strAuthCredentials = "Bearer " & strOpenAI_APIKey
    requestJSON = OpenAI_Model2JSON(strInputTextUser, strOpenAI_Model)
    
    On Error GoTo OpenAI_ChatCompletion
    
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    
    'Open and send the HTTP request
    Call httpRequest.Open("POST", strURL, False)
    Call httpRequest.SetRequestHeader("Content-Type", "application/json")
    Call httpRequest.SetRequestHeader("Authorization", strAuthCredentials)
    Call httpRequest.Send(requestJSON)
    
    requestReadyState = httpRequest.readyState

    'Check if the request is successful
    If requestReadyState = 1 Then
        ThisWorkbook.Sheets("VBA").TextBox4 = "Error: httpRequest.readyState=1"
        OpenAI_ChatCompletion = False
    ElseIf requestReadyState = 4 Then
        requestStatus = httpRequest.Status
        completion = httpRequest.responseText
        ThisWorkbook.Sheets("VBA").TextBox4 = completion
        ThisWorkbook.Sheets("VBA").TextBox11 = requestStatus
        If requestStatus = 200 Then
            ThisWorkbook.Sheets("VBA").TextBox10 = "General checks: OK"
            
            'Parse the JSON response
            RESPONSE = OpenAI_Decode_JSON_completion(completion, strOpenAI_Model)
            tokens = OpenAI_Decode_JSON_completion_tokens(completion)
            pricing = CInt(tokens) * pricing_per_token
            ThisWorkbook.Sheets("VBA").TextBox3 = RESPONSE
            ThisWorkbook.Sheets("VBA").TextBox15 = tokens & " tokens"
            ThisWorkbook.Sheets("VBA").TextBox16 = "$ " & CStr(pricing)
            OpenAI_ChatCompletion = True
        Else
            'Parse the JSON response
            RESPONSE = OpenAI_Decode_JSON_ChatGPTError(completion)
            ThisWorkbook.Sheets("VBA").TextBox3 = RESPONSE
            OpenAI_ChatCompletion = False
        End If
    Else
        ThisWorkbook.Sheets("VBA").TextBox4 = "Error: httpRequest.readyState not defined"
        OpenAI_ChatCompletion = False
    End If
    
    On Error GoTo 0
    Exit Function
        
OpenAI_ChatCompletion:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure OpenAI_ChatCompletion, line " & Erl & "."
     
End Function
'-----------------------------------------------------------------------------------------------
Private Function OpenAI_Pricing(ByVal strOpenAI_Model As String) As Single

    Dim pricing_per_token As Single
    
    'Pricing pour 1000 tokens (Not updated in real time. Please check the website)

    If (strOpenAI_Model = "gpt-3.5-turbo") Then
        pricing_per_token = 0.002
    ElseIf (strOpenAI_Model = "gpt-4") Then
        pricing_per_token = 0.06
    ElseIf (strOpenAI_Model = "gpt-4-8k") Then 'prompt
        pricing_per_token = 0.03
     ElseIf (strOpenAI_Model = "gpt-4-32k") Then 'completion
        pricing_per_token = 0.06
    ElseIf (strOpenAI_Model = "ada") Then
        pricing_per_token = 0.0004
    ElseIf (strOpenAI_Model = "davinci" Or strOpenAI_Model = "test-davinci-003" Or strOpenAI_Model = "test-davinci-002") Then
        pricing_per_token = 0.02
    ElseIf (strOpenAI_Model = "curie") Then
        pricing_per_token = 0.002
    ElseIf (strOpenAI_Model = "babbage") Then
        pricing_per_token = 0.0005
    Else
        pricing_per_token = 0
    End If

    OpenAI_Pricing = pricing_per_token / 1000
    
End Function
'-----------------------------------------------------------------------------------------------
Public Function OpenAI_Model2JSON(ByVal strInputText As String, ByVal strOpenAI_Model As String) As String

    Dim strMessages As String
    Dim requestBody As String
    Dim strTextSystem As String
    Dim strModel As String
    
    strModel = ThisWorkbook.Sheets("VBA").ComboBox1
    
    'Const RESPONSE_FORMAT As String = "{""type"" : ""json_object""}"
    Const MAX_TOKENS As Long = 1024 'Defaults to 16. The maximum number of tokens that can be generated in the completion. The token count of your prompt plus max_tokens cannot exceed the model's context length.
    Const TEMPERATURE As Single = 1 'Defaults to 1. What sampling temperature to use, between 0 and 2. Higher values like 0.8 will make the output more random, while lower values like 0.2 will make it more focused and deterministic. We generally recommend altering this or top_p but not both.
    Const TOP_P As Single = 1 'Defaults to 1. An alternative to sampling with temperature, called nucleus sampling, where the model considers the results of the tokens with top_p probability mass. So 0.1 means only the tokens comprising the top 10% probability mass are considered. We generally recommend altering this or temperature but not both.
    Const N As Integer = 1 'Defaults to 1. How many completions to generate for each prompt.
     
    strTextSystem = ThisWorkbook.Sheets("VBA").TextBox13
    
    If (strOpenAI_Model = "gpt-3.5-turbo-instruct" Or strOpenAI_Model = "babbage-002" Or strOpenAI_Model = "davinci-002") Then
        requestBody = "{" & _
                    """model"": """ & strModel & """" & _
                    "," & _
                    """max_tokens"": " & MAX_TOKENS & _
                    "," & _
                    """temperature"": " & TEMPERATURE & _
                    "," & _
                    """top_p"": " & TOP_P & _
                    "," & _
                    """n"": " & N & _
                    "," & _
                    """prompt"": """ & strInputText & """" & _
                    "}"
    ElseIf (Left(strOpenAI_Model, 4) = "gpt-") Then
        strMessages = "[]"
        strMessages = OpenAI_AppendMessage(strMessages, "system", strTextSystem)
        strMessages = OpenAI_AppendMessage(strMessages, "user", strInputText)
        
        requestBody = "{" & _
            """model"": """ & strModel & """" & _
            "," & _
            """max_tokens"": " & MAX_TOKENS & _
            "," & _
            """temperature"": " & TEMPERATURE & _
            "," & _
            """top_p"": " & TOP_P & _
            "," & _
            """n"": " & N & _
            "," & _
            """messages"": " & strMessages & "" & _
            "}"
    End If
        
    OpenAI_Model2JSON = requestBody

End Function
'-----------------------------------------------------------------------------------------------
Public Function OpenAI_AppendMessage(ByVal strMessages As String, ByVal strRole As String, ByVal strInputText As String) As String

    Dim strJSON As String
    Dim lenMessages As Long

    strJSON = OpenAI_InputRole2JSON(strRole, strInputText)

    lenMessages = Len(strMessages)
    strMessages = Left(strMessages, lenMessages - 1)
    
    If lenMessages = 2 Then
        strMessages = strMessages & strJSON & "]"
    Else
        strMessages = strMessages & "," & strJSON & "]"
    End If
    
    OpenAI_AppendMessage = strMessages
    
End Function
'-----------------------------------------------------------------------------------------------
Public Function OpenAI_InputRole2JSON(ByVal strRole As String, ByVal strText As String) As String

    Dim strJSON As String
    Dim strContent As String
    
    strContent = Replace(strText, Chr(34), Chr(39))
    
    strJSON = "{" & _
        """role"": """ & strRole & """," & _
        """content"": """ & strContent & """" & _
        "}"

    OpenAI_InputRole2JSON = strJSON
    
End Function
'-----------------------------------------------------------------------------------------------
Private Function OpenAI_Decode_JSON_completion(ByVal completion As String, ByVal strOpenAI_Model As String) As String

    On Error GoTo OpenAI_Decode_JSON_Error
    
    Dim json1 As Object
    Dim json2 As Object
    Dim json3 As Object
    Dim textElement As String
    
    'Call the 'JsonConverter' VBA module
    If (strOpenAI_Model = "gpt-3.5-turbo-instruct" Or strOpenAI_Model = "babbage-002" Or strOpenAI_Model = "davinci-002") Then
        Set json1 = JsonConverter.ParseJson(completion)
        Set json2 = json1("choices")(1)
        textElement = json2("text")
    ElseIf (Left(strOpenAI_Model, 4) = "gpt-") Then
        Set json1 = JsonConverter.ParseJson(completion)
        Set json2 = json1("choices")(1)
        Set json3 = json2("message")
        textElement = json3("content")
    End If
    
    OpenAI_Decode_JSON_completion = textElement
    
    On Error GoTo 0
    Exit Function

OpenAI_Decode_JSON_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure OpenAI_Decode_JSON_completion, line " & Erl & "."

End Function
'-----------------------------------------------------------------------------------------------
Private Function OpenAI_Decode_JSON_completion_tokens(ByVal completion As String) As String

    On Error GoTo OpenAI_Decode_JSON_Error
    
    Dim json1 As Object
    Dim json2 As Object
    Dim textElement As String
    
    'Call the 'JsonConverter' VBA module
    Set json1 = JsonConverter.ParseJson(completion)
    Set json2 = json1("usage")
    textElement = json2("total_tokens")
    
    OpenAI_Decode_JSON_completion_tokens = textElement
    
    On Error GoTo 0
    Exit Function

OpenAI_Decode_JSON_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure OpenAI_Decode_JSON_completion_tokens, line " & Erl & "."

End Function
'-----------------------------------------------------------------------------------------------
Private Function OpenAI_Decode_JSON_ChatGPTError(ByVal completion As String) As String

    On Error GoTo OpenAI_Decode_JSON_ChatGPTError_Error
    
    Dim json1 As Object
    Dim json2 As Object
    Dim textElement As String
    
    'Call the 'JsonConverter' VBA module
    Set json1 = JsonConverter.ParseJson(completion)
    Set json2 = json1("error")
    textElement = json2("message")

    OpenAI_Decode_JSON_ChatGPTError = textElement
    
    On Error GoTo 0
    Exit Function

OpenAI_Decode_JSON_ChatGPTError_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure OpenAI_Decode_JSON_ChatGPTError, line " & Erl & "."

End Function
'-----------------------------------------------------------------------------------------------
Public Function OpenAI_Test_URL(ByVal strOpenAI_URL As String, ByVal strOpenAI_APIKey As String, ByVal strOpenAI_Model As String) As Boolean

    Dim httpRequest As MSXML2.XMLHTTP60
    Dim strAuthCredentials As String
    Dim requestReadyState As Long
    Dim requestStatus As Long
    Dim completion As String
    Dim strURL As String
   
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    
    strURL = strOpenAI_URL & "/v1/models/" & strOpenAI_Model
    strAuthCredentials = "Bearer " & strOpenAI_APIKey
    
    'Verification de la connexion a l'URL
    Call httpRequest.Open("GET", strURL, False)
    Call httpRequest.SetRequestHeader("Content-Type", "application/json")
    Call httpRequest.SetRequestHeader("Authorization", strAuthCredentials)
    Call httpRequest.Send
    
    
    'Application.Wait Now + TimeValue( 0:00:03 )
    
    'Do While httpRequest.readyState <> 1
        'If time_chrono >= TIMEOUT Exit Do
        'DoEvents
    'Loop
    
    requestReadyState = httpRequest.readyState
    
    'Check if the request is successful
    If requestReadyState = 1 Then
        ThisWorkbook.Sheets("VBA").TextBox10 = "Error: httpRequest.readyState=" & CStr(requestReadyState)
        OpenAI_Test_URL = False
    ElseIf requestReadyState = 4 Then
        requestStatus = httpRequest.Status
        completion = httpRequest.responseText
        ThisWorkbook.Sheets("VBA").TextBox11 = requestStatus
        If requestStatus = 200 Then
            ThisWorkbook.Sheets("VBA").TextBox10 = completion
            OpenAI_Test_URL = True
        ElseIf requestStatus = 400 Then
            ThisWorkbook.Sheets("VBA").TextBox10 = completion
            OpenAI_Test_URL = True
        ElseIf requestStatus = 405 Then
            ThisWorkbook.Sheets("VBA").TextBox10 = completion
            OpenAI_Test_URL = False
        Else
            ThisWorkbook.Sheets("VBA").TextBox10 = completion
            OpenAI_Test_URL = False
        End If
    Else
        ThisWorkbook.Sheets("VBA").TextBox10 = "Error: httpRequest.readyState not defined: " & CStr(requestReadyState)
        OpenAI_Test_URL = False
    End If
    
End Function
'-----------------------------------------------------------------------------------------------
Public Function OpenAI_GetModelData(ByVal strOpenAI_URL As String, ByVal strOpenAI_APIKey As String, ByVal strOpenAI_Model As String) As String

    Dim httpRequest As MSXML2.XMLHTTP60
    Dim strAuthCredentials As String
    Dim requestReadyState As Long
    Dim requestStatus As Long
    Dim completion As String
    Dim strURL As String
    
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    
    strURL = strOpenAI_URL & "/v1/models/" & strOpenAI_Model
    strAuthCredentials = "Bearer " & strOpenAI_APIKey
    
    
    'Verification de la connexion a l'URL
    Call httpRequest.Open("GET", strURL, False)
    Call httpRequest.SetRequestHeader("Content-Type", "application/json")
    Call httpRequest.SetRequestHeader("Authorization", strAuthCredentials)
    Call httpRequest.Send
    
    requestReadyState = httpRequest.readyState
    
    If requestReadyState = 1 Then
        OpenAI_GetModelData = ""
    ElseIf requestReadyState = 4 Then
        requestStatus = httpRequest.Status
        completion = httpRequest.responseText
        ThisWorkbook.Sheets("VBA").TextBox11 = requestStatus
        OpenAI_GetModelData = completion
    End If

End Function
'-----------------------------------------------------------------------------------------------
Private Function OpenAI_Errors(ByVal strErrorCode As String, ByVal strCompletion As String) As String

    Dim strErrorMessage As String
    Dim json1 As Object
    
    strErrorMessage = ""
    
    'Call the 'JsonConverter' VBA module
    Set json1 = JsonConverter.ParseJson(strCompletion)
    
    If strErrorCode = 400 Then
        '1 - you must provide a model parameter
    ElseIf strErrorCode = 401 Then
        '1- Invalid Authentication
        '2- You didn't provide qn API key
        '3- Incorrect API key provided
        '4- You must be a member of an organization to use the API
    ElseIf strErrorCode = 429 Then
        '1- Rate limit reached for requests
        '2- You exceeded your current quota,please check your plan and billing details
        '3- The engine is currently overloaded,please try again later
    ElseIf strErrorCode = 500 Then
        '1 - The server had an error while processing your request
    End If
    
    OpenAI_Errors = strErrorMessage

End Function
