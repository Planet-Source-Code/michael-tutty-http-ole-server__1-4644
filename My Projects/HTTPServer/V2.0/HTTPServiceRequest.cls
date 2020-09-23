VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HTTPServiceRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private WithEvents mSock As SocketConnection
Attribute mSock.VB_VarHelpID = -1
Private mParent As HTTPServer
Private mstrData As String

Public Method As String
Public URI As String
Public HTTPVersion As String
Public Headers As HTTPHeaderList
Public Parameters As HTTPParameterList
Public Index As Long

Const GET_DELIMITER = "?"
Const HEADER_DELIMITER = vbNewLine & vbNewLine
Friend Function Initialize(Conn As SocketConnection, Parent As HTTPServer) As Boolean

  Set mSock = Conn
  Set mParent = Parent
  
  Initialize = (Not (Conn Is Nothing)) And (Not (mParent Is Nothing))

End Function


Private Function ParseElements(ByVal Data As String) As Boolean

  Dim FirstLine As String
  Dim URITemp As String
  Dim HeaderTemp As String
  
  If (Right(Data, 2) <> (vbCr & vbCr)) And (Right(Data, 2) <> vbNewLine) Then
    Exit Function
  End If
  
  Set Parameters = New HTTPParameterList
  Set Headers = New HTTPHeaderList

  FirstLine = SplitFrom(Data, vbCr)
  
  Method = Trim(SplitFrom(FirstLine, " "))
  URITemp = Trim(SplitFrom(FirstLine, " "))
  HTTPVersion = Trim(FirstLine)
  
  URI = SplitFrom(URITemp, GET_DELIMITER)

  If Len(URITemp) Then
    ' Process GET commands
    Parameters.AddItems URITemp
  End If
  
  HeaderTemp = SplitFrom(Data, vbCr & vbCr)
  If Len(HeaderTemp) Then
    Headers.AddItems HeaderTemp
  Else
    Headers.AddItems Data
    Data = ""
  End If
  
  If Len(Data) Then
    Parameters.AddItems Data
  End If
  
  ParseElements = True

End Function


Public Property Get RawRequest() As String

  RawRequest = mstrData

End Property

Private Sub SendDebugResponse(Optional Data As String = "")

  Dim sTemp As String
  Dim sContent As String
  
  If Data = "" Then
    sContent = "<HTML><BODY>Here's where the response would be</BODY></HTML>"
  Else
    sContent = Data
  End If
    
  sTemp = "HTTP/1.0 200 OK" & vbNewLine
  sTemp = sTemp & "Content-type: text/html" & vbNewLine
  sTemp = sTemp & "Content-length: " & Len(sContent) & vbNewLine
  sTemp = sTemp & vbNewLine
  
  sTemp = sTemp & sContent

  mSock.SendData sTemp

End Sub

Private Sub SendSocketClose()

  On Error Resume Next
  DoEvents
  mSock.SendData Chr(0)
  mSock.SendData Chr(0)
  mSock.SendData Chr(0)

End Sub

Friend Property Get Socket() As SocketConnection

  Set Socket = mSock

End Property

Public Sub WriteResponse(ResponseCode As String, HeaderData As DelimitedString, Data As String)

  Dim RetString As String
  
  RetString = Me.HTTPVersion & " " & ResponseCode & vbCrLf & HeaderData.Value & vbCrLf & vbCrLf & Data & vbCrLf & vbCrLf

  mSock.SendData RetString
  Set mSock = Nothing

End Sub

Private Sub Class_Terminate()
  
  Set mSock = Nothing
  Set Parameters = Nothing

End Sub


Private Sub mSock_ConnectionClosed()

  On Error Resume Next
  Set mSock = Nothing
  mParent.Dispose Me.Index
  Set mParent = Nothing

End Sub


Private Sub mSock_DataReceived(Data As Variant, Count As Long)

  mstrData = mstrData & Data

  If ParseElements(mstrData) Then
    mParent.GetResponse Me
    mstrData = ""
  End If

End Sub
Private Sub mSock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

  WriteStatus "Error: " & Number & ", " & Description
  Call mSock_ConnectionClosed

End Sub

