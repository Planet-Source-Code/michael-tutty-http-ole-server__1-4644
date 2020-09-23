VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SocketConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event ConnectionClosed()
Event Connect()
Event Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Event DataReceived(Data, Count As Long)
Event SendComplete()
Event SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)

Private WithEvents mSocket As MSWinsockLib.Winsock
Attribute mSocket.VB_VarHelpID = -1
Private mSocketFactory As ISocketFactory
Private mLog As ILogProvider
Friend Function Initialize(RequestID As Long, SocketFactory As ISocketFactory, LogProvider As ILogProvider) As Boolean

  'On Error Resume Next
  
  Set mSocketFactory = SocketFactory
  Set mSocket = mSocketFactory.GetConnectionSocket()
  Set mLog = LogProvider
  
  If mSocket.State <> sckClosed Then
    mSocket.Close
    If WaitForState(mSocket, sckClosed) = False Then
      Exit Function
    End If
  End If
  
  mSocket.Accept RequestID
  
  If WaitForState(mSocket, sckConnected) = False Then
    Exit Function
  End If
  
  ' Not testing for mSocketFactory, since if mSocket is valid, it
  ' must have come from a valid SocketFactory
  Initialize = (Not (mSocket Is Nothing)) And (Not (mLog Is Nothing))

End Function

Private Sub Class_Terminate()

  Dim CtlName As String

  On Error Resume Next
  mLog.WriteLog rmtConnect, "SocketConnection (" & mSocket.SocketHandle & ") terminating"

  If mSocket.State <> sckClosed Then
    mSocket.Close
    If WaitForState(mSocket, sckClosed) = False Then
      mLog.WriteLog rmtError, LoadResString(SS_ERRCODE_STATETIMEOUT)
    End If
  End If
  
  CtlName = mSocket.Name
  Set mSocket = Nothing
  
  If mSocketFactory.ReleaseConnectionSocket(CtlName) = False Then
    mLog.WriteLog rmtError, LoadResString(SS_ERRCODE_TERMCONN)
  End If
  
  Set mSocketFactory = Nothing
  Set mLog = Nothing

End Sub


Private Sub mSocket_Close()

  RaiseEvent ConnectionClosed
  mLog.WriteLog rmtConnect, "Connection Closed"

End Sub

Private Sub mSocket_Connect()

  RaiseEvent Connect
  mLog.WriteLog rmtConnect, "Connection Opened"

End Sub


Private Sub mSocket_DataArrival(ByVal bytesTotal As Long)

  Dim sTemp

  mSocket.GetData sTemp, vbString
  RaiseEvent DataReceived(sTemp, bytesTotal)
  mLog.WriteLog rmtData, "Received " & bytesTotal & " bytes"

End Sub
Private Sub mSocket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

  RaiseEvent Error(Number, Description, Scode, Source, HelpFile, HelpContext, CancelDisplay)
  mLog.WriteLog rmtError, "Error: " & Number & ", " & Description & ": SCode=" & Scode & ", " & "Source=" & Source

End Sub
Private Sub mSocket_SendComplete()

  RaiseEvent SendComplete
  mLog.WriteLog rmtProgress, "Send Complete"

End Sub
Private Sub mSocket_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)

  RaiseEvent SendProgress(bytesSent, bytesRemaining)
  mLog.WriteLog rmtProgress, "Socket send progress: " & bytesSent & ", " & bytesRemaining

End Sub

Public Sub SendData(Data)
  
  Debug.Assert WaitForState(mSocket, sckConnected)
  mSocket.SendData CStr(Data)
  DoEvents
  Debug.Assert WaitForState(mSocket, sckConnected)
  
  On Error Resume Next
  mLog.WriteLog rmtData, "Sent " & Len(Data) & " bytes"

End Sub

