Attribute VB_Name = "Main"
Sub test_Oa()
    
    Dim Ou As New FlowOauth
    Dim client As String
    Dim token As String
    Dim apiKey As String
    
    apiKey = ThisWorkbook.Path & "\credentials\api_key.json"
    token = ThisWorkbook.Path & "\credentials\token.json"
    client = ThisWorkbook.Path & "\credentials\client_secret.json"
    
    With Ou
        .webBrowser = "brave.exe"
        .InitializeFlow client, token, apiKey, OU_SCOPE_DRIVE_READONLY
        Debug.Print "API KEY"; " -- "; .GetApiKey
        Debug.Print "TOKEN ACCESS "; " -- "; .GetTokenAccess
        Debug.Print vbCrLf
    End With

End Sub

