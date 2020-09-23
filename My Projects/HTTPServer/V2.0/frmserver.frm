VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServer 
   Caption         =   "Socket Server"
   ClientHeight    =   4575
   ClientLeft      =   3075
   ClientTop       =   2715
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   6405
   Begin VB.ListBox lstLog 
      Height          =   3765
      IntegralHeight  =   0   'False
      Left            =   30
      TabIndex        =   0
      Top             =   600
      Width           =   6315
   End
   Begin MSWinsockLib.Winsock wsListener 
      Left            =   30
      Top             =   30
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wsConnection 
      Index           =   0
      Left            =   510
      Top             =   30
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuLog 
      Caption         =   "&Log"
      Begin VB.Menu mnuLogLevelTop 
         Caption         =   "&Level"
         Begin VB.Menu mnuLogLevel 
            Caption         =   "&None"
            Index           =   0
         End
         Begin VB.Menu mnuLogLevel 
            Caption         =   "&Some"
            Index           =   1
         End
         Begin VB.Menu mnuLogLevel 
            Caption         =   "&Connections"
            Index           =   2
         End
         Begin VB.Menu mnuLogLevel 
            Caption         =   "&Data"
            Index           =   3
         End
         Begin VB.Menu mnuLogLevel 
            Caption         =   "&Transfer Progress"
            Index           =   4
         End
      End
      Begin VB.Menu mnuLogClear 
         Caption         =   "&Clear"
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mIndex As Integer
Private mLogLevel As LogLevels
Private mbEnabled As Boolean

Implements ISocketFactory
Implements ILogProvider

Private Const SOCKET_NAME_PREFIX = "wsSocket"
Private Function GetSocketControlName(Index As Integer) As String

  GetSocketControlName = SOCKET_NAME_PREFIX & Index

End Function

Friend Function Initialize() As Boolean

  Initialize = True

End Function

Private Sub SetLogLevel(NewLevel As LogLevels)

  mLogLevel = NewLevel

End Sub

Private Sub Form_Resize()

  lstLog.Move 0, 0, ScaleWidth, ScaleHeight

End Sub


Private Property Let ILogProvider_LogLevel(RHS As LogLevels)

  SetLogLevel RHS

End Property

Private Property Get ILogProvider_LogLevel() As LogLevels

  ILogProvider_LogLevel = mLogLevel

End Property

Private Sub ILogProvider_WriteLog(EntryType As LogLevels, LogEntry As String)

  If EntryType <= mLogLevel Then
    lstLog.AddItem "[" & Now & "] - " & LogEntry
  End If

End Sub

Private Function GetConnectionSocketOld() As MSWinsockLib.Winsock

  Dim objConn As MSWinsockLib.Winsock
  Dim iLoop As Long
  
  ' Find the first available (not loaded) control instance
  For iLoop = 1 To wsConnection.UBound
    On Error Resume Next
    If (wsConnection(iLoop) Is Nothing) Then
      Exit For
    End If
  Next iLoop
  
  Load wsConnection(iLoop)
  'Set ISocketFactory_GetConnectionSocket = wsConnection(iLoop)

End Function

Private Function ISocketFactory_GetConnectionSocket() As MSWinsockLib.IMSWinsockControl

  mIndex = mIndex + 1
  
  Set ISocketFactory_GetConnectionSocket = Me.Controls.Add("MSWinsock.Winsock.1", GetSocketControlName(mIndex), Me)

End Function

Private Function ISocketFactory_GetFactorySocket() As MSWinsockLib.IMSWinsockControl

  Set ISocketFactory_GetFactorySocket = wsListener

End Function


Private Function ISocketFactory_ReleaseConnectionSocket(ControlName As String) As Boolean

'  On Error Resume Next
  Me.Controls.Remove ControlName

  ISocketFactory_ReleaseConnectionSocket = (Err.Number = 0)

End Function


Private Sub mnuFileExit_Click()
  
  ILogProvider_WriteLog rmtNone, "Shutting down..."

  Unload Me

End Sub


Private Sub mnuLogClear_Click()

  lstLog.Clear

End Sub


Private Sub mnuLogLevel_Click(Index As Integer)

  mLogLevel = Index

End Sub

