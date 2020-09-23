VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmTestFrame 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Test Frame"
   ClientHeight    =   5490
   ClientLeft      =   2535
   ClientTop       =   2190
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   Begin InetCtlsObjects.Inet inetTest 
      Left            =   1470
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      RequestTimeout  =   15
   End
   Begin VB.TextBox txtLog 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   4875
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   600
      Width           =   6915
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Start Tests"
      Height          =   495
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1275
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   1410
      TabIndex        =   2
      Top             =   150
      UseMnemonic     =   0   'False
      Width           =   5385
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmTestFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mobjSrv As HTTPServer
Attribute mobjSrv.VB_VarHelpID = -1
Private Sub ClearLog()

  txtLog.Text = ""

End Sub

Public Sub LogItem(LogData As String)

  txtLog.Text = txtLog.Text & LogData & vbNewLine
  txtLog.SelStart = Len(txtLog.Text)

End Sub

Private Sub RunTests()

  If Test1() = False Then
    Status = "Failed Test #1"
    Exit Sub
  End If
  
  LogItem ""
  
  If Test2() = False Then
    Status = "Failed Test #2"
    Exit Sub
  End If

  LogItem ""

  Set mobjSrv = Nothing

  Status = "All tests completed."
  LogItem ""
  LogItem "All tests completed."

End Sub
Public Property Let Status(newStatus As String)

  lblStatus.Caption = newStatus
  lblStatus.Refresh

End Property
Private Function Test1() As Boolean

  Status = "Running Test #1(Init)"
  LogItem "Starting Test #1 at " & Now
  
  Dim objSrv As HTTPServer
  
  LogItem "Creating HTTPServer"
  Set objSrv = New HTTPServer
  
  LogItem "Initializing HTTPServer to port 7777"
  If objSrv.Initialize(7777) = False Then
    LogItem "Initialize failed with port 7777"
    Exit Function
  End If
  
  Set objSrv = Nothing
  
  LogItem "Test 1 Succeeded."
  Test1 = True

End Function
Private Function Test2() As Boolean

  Status = "Running Test #2 HTTP Request"
  LogItem "Starting Test #2 at " & Now
  
  Dim sURL As String
  Dim sRet As String
  
  LogItem "Creating HTTPServer"
  Set mobjSrv = New HTTPServer
  
  LogItem "Initializing HTTPServer to port 1234"
  On Error Resume Next
  If mobjSrv.Initialize(1234) = False Then
    LogItem "Initialize failed with port 1234"
    Exit Function
  End If
  
  sURL = "http://localhost:1234/test/testpage.htm"
  LogItem "Opening test page from server at: " & sURL
  sRet = inetTest.OpenURL(sURL, icString)
  
  If sRet <> "" Then
    LogItem "Return Data = " & vbNewLine & sRet
    LogItem "Test 2 Succeeded."
    Test2 = True
  Else
    LogItem "Request to " & sURL & " returned no data."
    LogItem "Test 2 Failed."
    Test2 = False
  End If

End Function

Private Sub cmdTest_Click()

  Set mobjSrv = Nothing
  ClearLog
  RunTests

End Sub

Private Sub Form_Unload(Cancel As Integer)

  Set mobjSrv = Nothing

End Sub


Private Sub mobjSrv_OnRequest(Request As HTTPServerLib.HTTPServiceRequest)

  Dim HeaderData As New DelimitedString
  Dim ResponseData As New DelimitedString

  LogItem "Received a request for " & Request.URI
  LogItem "Method = " & Request.Method
  LogItem "Raw data = " & Request.RawRequest
  
  HeaderData.Add "Connection: Close"
  With ResponseData
    .Add "<html>"
    .Add "<body>"
    .Add "<h1>You requested " & Request.URI & "</h1>"
    .Add "</body></html>"
  End With
  Request.WriteResponse "200 OK", HeaderData, ResponseData.Value

End Sub


