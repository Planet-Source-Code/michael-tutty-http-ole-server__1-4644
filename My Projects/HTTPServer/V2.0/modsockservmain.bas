Attribute VB_Name = "modSockServMain"
Option Explicit

Public Const SS_ERROR_SOURCE = "SoketServer"

Public Const SS_ERRCODE_INITFAIL = 1000
Public Const SS_ERRCODE_PORTCHANGE = 1001
Public Const SS_ERRCODE_BADPORT = 1002
Public Const SS_ERRCODE_INITCONN = 1003
Public Const SS_ERRCODE_TERMCONN = 1004
Public Const SS_ERRCODE_STATETIMEOUT = 1005

Public Enum LogLevels
  rmtNone
  rmtError
  rmtConnect
  rmtData
  rmtProgress
End Enum

Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Function WaitForState(Sock As MSWinsockLib.Winsock, State As MSWinsockLib.StateConstants, Optional Timeout As Long = 1000) As Boolean

  Dim TimeoutTick As Long
  
  TimeoutTick = GetTickCount + Timeout
  
  Do Until Sock.State = State
    DoEvents
    If GetTickCount > TimeoutTick Then
      Exit Function
    End If
  Loop
  
  WaitForState = True

End Function


