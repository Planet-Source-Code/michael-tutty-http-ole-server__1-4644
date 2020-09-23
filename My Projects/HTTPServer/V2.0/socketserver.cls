VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SocketServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event ConnectionRequest(ID As Long, Accept As Boolean)
Event NewConnection(Conn As SocketConnection)

Private mbEnabled As Boolean

Private WithEvents mSocket As MSWinsockLib.Winsock
Attribute mSocket.VB_VarHelpID = -1
Private mSocketFactory As ISocketFactory
Private mLog As ILogProvider

Public Property Get LogProvider() As ILogProvider

  Set LogProvider = mLog

End Property

Public Property Get SocketFactory() As ISocketFactory

  Set SocketFactory = mSocketFactory

End Property

Private Sub Class_Terminate()

  Dim fTemp As frmServer
  
  Set fTemp = mSocketFactory
  
  Set mSocketFactory = Nothing
  Set mLog = Nothing
  
  Unload fTemp
  Set fTemp = Nothing

End Sub

Private Sub mSocket_ConnectionRequest(ByVal RequestID As Long)

  Dim NewConn As SocketConnection
  Dim bAccept As Boolean

  mLog.WriteLog rmtConnect, "Connection Requested, ID=" & RequestID
  
  RaiseEvent ConnectionRequest(RequestID, bAccept)
  If bAccept Then
    ' Enabled = False
    Set NewConn = New SocketConnection
    If NewConn.Initialize(RequestID, mSocketFactory, mLog) Then
      mLog.WriteLog rmtConnect, "Connection Accepted, ID=" & RequestID
      RaiseEvent NewConnection(NewConn)
    Else
      RaiseError SS_ERRCODE_INITCONN
    End If
    Set NewConn = Nothing
    ' Enabled = True
  End If

End Sub
Public Sub RaiseError(ErrNumber As Long)

  Dim sErr As String

  sErr = LoadResString(ErrNumber)
  Err.Raise vbObjectError + ErrNumber, SS_ERROR_SOURCE, sErr
  mLog.WriteLog rmtError, "Error: " & ErrNumber & " [" & sErr & "]"

End Sub

Private Sub mSocket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

  mLog.WriteLog rmtError, "Error: " & Number & ", " & Description & ": SCode=" & Scode & ", " & "Source=" & Source

End Sub
Private Sub OnEnabledChange(NewEnabled As Boolean)

  If mSocket.LocalPort < 1 Then
    RaiseError SS_ERRCODE_BADPORT
    Exit Sub
  End If
  
  If NewEnabled Then
    If mSocket.State <> sckClosed Then
      mSocket.Close
      If WaitForState(mSocket, sckClosed) = False Then
        RaiseError SS_ERRCODE_STATETIMEOUT
        Exit Sub
      End If
    End If
    mSocket.Bind mSocket.LocalPort
    mSocket.Listen
    If WaitForState(mSocket, sckListening) = False Then
      RaiseError SS_ERRCODE_STATETIMEOUT
      Exit Sub
    End If
    mbEnabled = True
  Else
    mSocket.Close
    If WaitForState(mSocket, sckClosed) = False Then
      RaiseError SS_ERRCODE_STATETIMEOUT
      Exit Sub
    End If
    mbEnabled = False
  End If

End Sub

Private Sub OnPortChange(NewPort As Long)

  If mbEnabled Then
    RaiseError SS_ERRCODE_PORTCHANGE
    Exit Sub
  End If

  mSocket.LocalPort = NewPort

End Sub

Public Property Let ServerPort(NewServerPort As Long)

  If NewServerPort <> mSocket.LocalPort Then
    OnPortChange NewServerPort
  End If

End Property

Public Property Get ServerPort() As Long

  ServerPort = mSocket.LocalPort

End Property

Public Property Get Enabled() As Boolean

  Enabled = mbEnabled

End Property

Public Property Let Enabled(ByVal NewEnabled As Boolean)

  If NewEnabled <> mbEnabled Then
    OnEnabledChange NewEnabled
  End If

End Property

Private Sub Class_Initialize()

  Dim fTemp As New frmServer
  
  Load fTemp
  If fTemp.Initialize() = False Then
    RaiseError SS_ERRCODE_INITFAIL
    Unload fTemp
  Else
    Set mLog = fTemp
    Set mSocketFactory = fTemp
    Set mSocket = mSocketFactory.GetFactorySocket()
    #If debugmode = 1 Then
      fTemp.Show
    #Else
      fTemp.Hide
    #End If
  End If
  
  Set fTemp = Nothing

End Sub



