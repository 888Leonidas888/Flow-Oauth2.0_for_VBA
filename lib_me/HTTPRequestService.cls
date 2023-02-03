VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HTTPRequestService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event forRequestHTTP(ByVal status As Integer)

Private Type tState
    allHeaders As String
    responseText As String
End Type

Private This As tState

Public Function HTTPrequest(ByVal methods As String, _
                            ByVal url As String, _
                            Optional body = Empty, _
                            Optional dicHeaders As Dictionary) As String
    
    Dim HTTP As New MSXML2.ServerXMLHTTP60
    
    With HTTP
    
        .Open methods, url, False
        
        If Not dicHeaders Is Nothing Then
            For Each key In dicHeaders.Keys
                .setRequestHeader key, dicHeaders(key)
            Next key
        End If
        
        .send body
        
        RaiseEvent forRequestHTTP(.status)
        
        With This
            .allHeaders = HTTP.getAllResponseHeaders
            .responseText = HTTP.responseText
        End With
         
        HTTPrequest = This.responseText
        
    End With
    
    Set HTTP = Nothing
    
End Function

