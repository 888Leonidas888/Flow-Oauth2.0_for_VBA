VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FlowOauth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'**************************************************************************
'>>>Desarrollado por: Jhony Escriba Bautista
'>>>Correo : jhonny14_1@hotmail.com
'>>>Año de lanzamiento 2023
'>>>Version de Oauth 2.0
'***************************************************************************
Private WithEvents HTTP As HTTPRequestService
Attribute HTTP.VB_VarHelpID = -1

Private Const APP_NAME As String = "Flujo de Oauth 2.0"
Private Const VERSION As String = "2.0.0.0"

Private Type tFo
    yourApiKey As String
    yourAccessToken As String
    client_id As String
    redirect_uri As String
    client_secret As String
    refresh_token As String
    status As Integer
    responseTxt As String
End Type

Public webBrowser As String

Private Const OAUTH2_END_POINT As String = "https://accounts.google.com/o/oauth2/v2/auth"
Private Const TOKEN_END_POINT As String = "https://oauth2.googleapis.com/token"
Private Const REVOKE_TOKEN_END_POINT As String = "https://oauth2.googleapis.com/revoke"

Private This As tFo

Public Function GetTokenAccess() As String
    GetTokenAccess = This.yourAccessToken
End Function

Public Function GetApiKey() As String
    GetApiKey = This.yourApiKey
End Function

Public Property Get Operation() As Variant
    Operation = This.status
End Property

Public Sub InitializeFlow(ByVal credentialsClient As String, _
                                ByVal credentialsToken As String, _
                                ByVal credentialsApiKey As String, _
                                ByVal scope As String)
                                
        Dim stringRedirect As String
        Dim ExcelObj As New Excel.Application
        Dim code As String
        Dim content As String
        
        Debug.Print APP_NAME; " "; Now(); " >>> "; "flow started"
        
        Rem lectura de api_key;client_id;redirect_uris;client_secret
        Call readCredentialsGoogleConsole(credentialsClient, credentialsApiKey)
        
        If ExistsFile(credentialsToken) Then
            Call verifyingToken(credentialsToken)
        Else
            Debug.Print APP_NAME; " "; Now(); " >>> "; "FILE NOT FOUND " & credentialsToken
            
        Rem #1.- Autenticar y autorizar a al APP;redirección del usuario a la pantalla de consentimiento
            stringRedirect = urlRedirect(scope)
            Shell "cmd /c start " & Me.webBrowser & " " & """" & stringRedirect & """", vbHide
'            Shell "cmd /c start chrome.exe " & """" & stringRedirect & """", vbHide

        Rem #2.- Obtener código para cambiar
            ExcelObj.Wait Now() + TimeValue("00:00:05")
            code = ExcelObj.InputBox( _
            "Ingrese el valor del parámetro ""code"" que aparece en la URL de su navegador", _
            APP_NAME, Type:=2)
            Set ExcelObj = Nothing
            
            If code = Empty Then Exit Sub
            
        Rem #3.- Cambiar el codígo para un token de acceso
            content = changeCodeForToken(code)
        
        Rem #4.- Guardar el token de acceso para su posterior uso.
            If Me.Operation = GO_SUCCESSFUL Then
                Dim currentDate As Double
                Dim currentDateField As String
                
                'agregamos la fecha y la hora menos 10 segundos en la que se escribe el token.
                currentDate = CDbl(Date) + CDbl(DateAdd("s", 3599 - 300, Time))
                currentDateField = ",""current_date"":" & currentDate & "}"
                content = Replace(content, "}", currentDateField)
                '----------------------------------------------------
                
                Call readTokenBeforeWrite(content)
                webBrowserHTTP.writeFile content, credentialsToken
                
                Debug.Print APP_NAME; " "; Now(); " >>> "; "new token generated"
            Else
                Debug.Print APP_NAME; " "; Now(); " >>> "; "error "; content
            End If
        End If
        
End Sub

Public Function RevokeToken(ByVal credentialsToken As String) As Boolean
    Dim body As String
    Dim headers As New Dictionary
    Dim dicToken As Dictionary
    Dim content As String
    
    If Not webBrowserHTTP.ExistsFile(credentialsToken) Then Exit Function
    
    content = webBrowserHTTP.readFile(credentialsToken)
    Set dicToken = JsonConverter.ParseJson(content)
    This.yourAccessToken = dicToken("access_token")
    
    body = "token=" & This.yourAccessToken
    headers.Add "Content-type", "application/x-www-form-urlencoded"
    HTTPFlowOauth HT_POST, REVOKE_TOKEN_END_POINT, body, headers
    
    If Me.Operation = GO_SUCCESSFUL Then
        Debug.Print APP_NAME; " "; Now(); " >>> "; "revoked token"
        RevokeToken = True
    Else
        Debug.Print APP_NAME; " "; Now(); " >>> "; "failed token revoke"
        RevokeToken = False
    End If
    Set dicToken = Nothing
    Set headers = Nothing
End Function
Private Sub readTokenBeforeWrite(ByVal content As String)
    Dim dicToken As Dictionary

    Set dicToken = JsonConverter.ParseJson(content)
    
    With This
        .yourAccessToken = dicToken("access_token")
    End With

    Set dicToken = Nothing
End Sub

Private Sub readCredentialsGoogleConsole(ByVal credentialsClient As String, ByVal credentialsApiKey As String)
        
    Dim dicApikey As Dictionary
    Dim dicClient As Dictionary
    Dim content As String
    Dim app As String
    
    If ExistsFile(credentialsApiKey) Then
        content = readFile(credentialsApiKey)
        Set dicApikey = JsonConverter.ParseJson(content)
        This.yourApiKey = dicApikey("your_api_key")
    End If
          
    If ExistsFile(credentialsClient) Then
        content = readFile(credentialsClient)
        Set dicClient = JsonConverter.ParseJson(content)
        With dicClient
            If .Exists("web") Then app = "web"
            If .Exists("installed") Then app = "installed"
        End With
        
        With This
            .client_id = dicClient(app)("client_id")
            .redirect_uri = dicClient(app)("redirect_uris")(1)
            .redirect_uri = codificationUrl(.redirect_uri)
            .client_secret = dicClient(app)("client_secret")
        End With
    End If
    
    Set dicClient = Nothing
    Set dicApikey = Nothing

End Sub
Private Function urlRedirect(ByVal scope As String) As String
        
        Dim response_type As String
        Dim access_type As String
        Dim include_granted_scopes As String
        Dim prompt As String
        Dim parametersPath As String
 
        response_type = "code"
        scope = codificationUrl(scope)
        access_type = "offline"
        include_granted_scopes = "true"
        prompt = "consent"
        
        parametersPath = formatString( _
        "?scope={0}&access_type={1}&include_granted_scopes={2}&response_type={3}&redirect_uri={4}&client_id={5}&prompt={6}", _
                        scope, _
                        access_type, _
                        include_granted_scopes, _
                        response_type, _
                        This.redirect_uri, _
                        This.client_id, prompt)
        
        urlRedirect = OAUTH2_END_POINT & parametersPath
            
End Function

Private Function changeCodeForToken(ByVal code As String) As String
    Rem esta funcion reemplaza a ChangeTokenUpDate
        
    Dim body As String
    Dim headers As New Dictionary
    
    body = formatString( _
            "code={0}&client_id={1}&client_secret={2}&redirect_uri={3}&grant_type={4}", _
            code, _
            This.client_id, _
            This.client_secret, _
            This.redirect_uri, _
            "authorization_code")
            
    headers.Add "Content-Type", "application/x-www-form-urlencoded"
    
    This.responseTxt = HTTPFlowOauth(HT_POST, TOKEN_END_POINT, body, headers)
    changeCodeForToken = This.responseTxt
    
    Set headers = Nothing
    
End Function

Private Sub verifyingToken(ByVal credentialsToken As String)
    
    Dim dicToken As Dictionary
    Dim error_ As String, error_description_ As String
    Dim rsp As Integer
    Dim content As String
    
    content = webBrowserHTTP.readFile(credentialsToken)
    Set dicToken = JsonConverter.ParseJson(content)
    
    Debug.Print APP_NAME; " "; Now(); " >>> "; "token verifying"
    
    If dicToken.Exists("error") Or dicToken.Exists("error_description") Then
        error_ = dicToken("error")
        error_description_ = dicToken("error_description")
        
        rsp = MsgBox(error_ & vbCrLf & error_description_ & vbCrLf & vbCrLf & _
            "Presione ""SI"" para eliminar el archivo con el token,esto generará un error en la respuesta de la API", _
            vbYesNo + vbDefaultButton1 + vbCritical, APP_NAME)
        
        If rsp = vbYes Then Kill credentialsToken
        
        Debug.Print APP_NAME; " "; Now(); " >>> "; error_; "  "; error_description_
        Rem "A DONDE ME VOY"
        Rem "A DONDE ME VOY"
        Rem "A DONDE ME VOY"
        Rem "A DONDE ME VOY"
        Rem "A DONDE ME VOY"
        Rem "A DONDE ME VOY"
        Rem "A DONDE ME VOY"
        Rem "A DONDE ME VOY"
        Set dicToken = Nothing
        Exit Sub
        
    ElseIf dicToken.Exists("access_token") And dicToken.Exists("token_type") Then
        
        Dim registeredDate As Double
        Dim refresh_token As String
        
        registeredDate = dicToken("current_date")
        
        If CDbl(Now()) > CDbl(registeredDate) Then
            This.refresh_token = dicToken("refresh_token")
            
            Debug.Print APP_NAME; " "; Now(); " >>> "; "requests update token"
            Call updateTokenAccess(credentialsToken)
        Else
            content = webBrowserHTTP.readFile(credentialsToken)
            Call readTokenBeforeWrite(content)
            Debug.Print APP_NAME; " "; Now(); " >>> "; "using token_access"
        End If
        Set dicToken = Nothing
    End If
    
    
End Sub

Private Sub updateTokenAccess(Optional credentialsToken As String)
    
    Dim body As String
    Dim headers As New Dictionary
    Dim content As String
    
    With This
    body = formatString( _
        "client_id={0}&client_secret={1}&refresh_token={2}&grant_type=refresh_token", _
        .client_id, .client_secret, .refresh_token)
    End With
    
    headers.Add "Content-Type", "application/x-www-form-urlencoded"
    content = HTTPFlowOauth(HT_POST, TOKEN_END_POINT, body, headers)
    
    If Me.Operation = GO_SUCCESSFUL Then
        Dim currentDate As Double
        Dim currentDateField As String
        
        Debug.Print APP_NAME; " "; Now(); " >>> "; "successfully updated token"
        
        'agregamas la fecha y la hora menos 5 segundos en la que se escribe el token.
        currentDate = CDbl(Date) + CDbl(DateAdd("s", 3599 - 300, Time))
        currentDateField = ",""current_date"":" & currentDate & "}"
        content = Replace(content, "}", currentDateField)
    
        'agregamos el refresh_token al contenido
        content = Replace(content, "{", "{" + vbCrLf + Chr(32) + Chr(32) + """refresh_token"":""" + This.refresh_token + """,")
        
        'leemos el token_access
        Call readTokenBeforeWrite(content)
        
    Else
        Debug.Print APP_NAME; " "; Now(); " >>> "; "failed update token"
        
    End If
    'guardamos respuesta exitosa o fallida
    webBrowserHTTP.writeFile content, credentialsToken
    Set headers = Nothing
    
End Sub

Private Function HTTPFlowOauth(ByVal methods As String, _
                                  ByVal url As String, _
                                  Optional body = Empty, _
                                  Optional dicHeaders As Dictionary) As String


    Set HTTP = New HTTPRequestService
    
    This.responseTxt = HTTP.HTTPrequest(methods, url, body, dicHeaders)
    
    HTTPFlowOauth = This.responseTxt
    
    Set HTTP = Nothing
    
End Function

Private Sub Class_Initialize()
    Me.webBrowser = """chrome.exe"""
End Sub

Private Sub HTTP_forRequestHTTP(ByVal status As Integer)
    This.status = status
End Sub
