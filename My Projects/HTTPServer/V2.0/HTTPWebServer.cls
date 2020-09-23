VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HTTPWebServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mServer As Server
Attribute mServer.VB_VarHelpID = -1

Friend Function Initialize(srv As Server) As Boolean

  Set mServer = srv
  
  Initialize = Not (mServer Is Nothing)

End Function


Private Sub Class_Terminate()

  Set mServer = Nothing

End Sub


Private Sub mServer_OnRequest(Request As HTTPServiceRequest)

  Dim HeaderData As New DelimitedString
  Dim ResponseData As String
  
  Const RESPONSE_DATA_TOP = "<html><body>Here's the page!<br>"
  Const RESPONSE_DATA_END = "</body></html>"
  
  DebugPrint Request, Request.RawRequest
  
  ResponseData = RESPONSE_DATA_TOP & "You requested " & Request.URI & "<br><hr>" & RESPONSE_DATA_END
  
  With HeaderData
    .Delimiter = vbNewLine
    .Add "Content-type: text/html"
    .Add "Content-length: " & Len(ResponseData)
  End With
  
  Request.WriteResponse "200 OK", HeaderData, ResponseData

End Sub


