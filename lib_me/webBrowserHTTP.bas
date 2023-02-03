Attribute VB_Name = "webBrowserHTTP"
Option Explicit

Public Function formatString(ByVal text As String, ParamArray values()) As String
    
    Dim i As Integer
    
    For i = LBound(values) To UBound(values)
        text = Replace(text, "{" & i & "}", values(i))
    Next i
    
    formatString = text
    
End Function

Public Function readFile(ByVal pathFile As String) As String
    
    Dim fso As New Scripting.FileSystemObject
    Dim t As TextStream
    Dim content As String
    
    If fso.FileExists(pathFile) Then
        Set t = fso.OpenTextFile(pathFile, ForReading)
        content = t.ReadAll
        t.Close
        readFile = content
    Else
        Err.Raise 1 + vbObjectError, _
        , _
        "No se encontro el archivo en esta ruta " + vbCrLf + vbCrLf + pathFile
        
    End If
    
    Set fso = Nothing
    Set t = Nothing
    
End Function

Public Function writeFile(ByVal content As String, Optional pathTarget) As Boolean
    
    Dim fso As New Scripting.FileSystemObject
    Dim t As TextStream
    Dim userProfile As String
    
    On Error GoTo Cath
    
    userProfile = Environ$("UserProfile")
    
    If IsMissing(pathTarget) Then: pathTarget = userProfile & "\content.txt"
    
    Set t = fso.CreateTextFile(pathTarget, True)
    t.Write content
    t.Close
    
    writeFile = True
    Exit Function
    
Cath:

    writeFile = False
    
End Function

Public Function ExistsFile(ByVal pathFile As String) As Boolean
    
    Dim fso As New Scripting.FileSystemObject
    ExistsFile = fso.FileExists(pathFile)
    Set fso = Nothing
    
End Function

Public Function generateString(Optional lenght = 8, Optional includeNumber = False) As String
    
    Dim dicUpper  As New Dictionary
    Dim dicLower As New Dictionary
    Dim dicNumbers As New Dictionary
    Dim randomString As String
    Dim character As Long
    Dim i As Integer
    
    'números del 48 al 57
    'letras mayúsculas 65 al 90
    'letras minúsculas 97 al 122
    
    For i = 48 To 57
        With dicNumbers
            .Add i, Empty
        End With
    Next i
    
    For i = 65 To 90
        With dicUpper
            .Add i, Empty
        End With
    Next i
    
    For i = 97 To 122
        With dicLower
            .Add i, Empty
        End With
    Next i
    
    
    Do While Len(randomString) <= lenght
        Randomize
        character = Int((122 - 48 + 1) * Rnd + 48)
        
         If (dicUpper.Exists(character) Or dicLower.Exists(character)) Or _
            (dicNumbers.Exists(character) And includeNumber) Then
            
            randomString = randomString + Chr(character)
            
        End If
    Loop
    
    Set dicUpper = Nothing
    Set dicLower = Nothing
    Set dicNumbers = Nothing
    
    generateString = randomString
    
End Function

Public Function codificationUrl(ByVal url As String) As String

    Dim dicCharacterSpecial As New Dictionary
    Dim key
    
    With dicCharacterSpecial
    Rem el primer Item debe ser el de porcentaje, debido a que al los demas valores reemplazados incluyen el porcentaje
        .Add "%", "%25"
        .Add " ", "%20"
        .Add "=", "%3D"
        .Add ",", "%2C"
        .Add """", "%22"
        .Add "<", "%3C"
        .Add ">", "%3E"
        .Add "#", "%23"
        .Add "|", "%7C"
        .Add "/", "%2F"
        .Add ":", "%3A"
        .Add "_", "%5F"
    End With
    
    For Each key In dicCharacterSpecial.Keys
        url = Replace(url, key, dicCharacterSpecial(key))
    Next key
    
    Set dicCharacterSpecial = Nothing
    
    codificationUrl = url
    
End Function

Function ConsoleShow(ByVal text As String, ByVal spaces As Integer) As String
    
    Dim strLength%, newLengthStr%
    Dim difSpaceLen%, i%
    
    strLength = Len(text)
    
    If spaces <= strLength Then
        text = Left$(text, spaces)
    Else
        difSpaceLen = spaces - strLength
        For i = 1 To difSpaceLen
            text = " " & text
        Next i
    End If
    
    ConsoleShow = "|" & text
    
End Function
Public Function boolToStr(Optional bool = False) As String
    
    Dim strBool As String
    
    strBool = "false"
    If bool Then strBool = "true"
    boolToStr = strBool
        
End Function



