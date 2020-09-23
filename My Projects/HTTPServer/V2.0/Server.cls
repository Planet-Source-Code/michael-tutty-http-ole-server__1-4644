VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HTTPServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private WithEvents mServer As SocketServer
Attribute mServer.VB_VarHelpID = -1
Private mCol As Collection

Event OnRequest(Request As HTTPServiceRequest)
Friend Sub Dispose(Index As Long)

  Dim iLoop As Long
  
  For iLoop = 1 To mCol.Count
    If mCol(iLoop).Index = Index Then
      mCol.Remove iLoop
      RemoveConnection
      Exit Sub
    End If
  Next iLoop
  
  Debug.Assert False

End Sub

Friend Sub GetResponse(Request As HTTPServiceRequest)

  RaiseEvent OnRequest(Request)

End Sub

Public Function Initialize(BindPort As Long) As Boolean

  On Error GoTo InitError

  Set mServer = New SocketServer
  
  mServer.ServerPort = BindPort
  
  mServer.Enabled = True
  
  Set mCol = New Collection
  
  Initialize = mServer.Enabled

InitError:
  Debug.Assert Err.Number = 0

End Function

Private Sub Class_Terminate()

  Dim iLoop As Long
  
  For iLoop = 1 To mCol.Count
    mCol.Remove iLoop
    RemoveConnection
  Next iLoop
  
  mServer.Enabled = False
  DoEvents
  Set mServer = Nothing

End Sub


Private Sub mServer_ConnectionRequest(ID As Long, Accept As Boolean)

  Accept = True

End Sub


Private Sub mServer_NewConnection(Conn As SocketConnection)

  Dim objNew As HTTPServiceRequest

  Static Index As Long

  Set objNew = New HTTPServiceRequest
  If objNew.Initialize(Conn, Me) Then
    Index = Index + 1
    objNew.Index = Index
    mCol.Add objNew
    AddConnection
  End If

  Set objNew = Nothing

End Sub

