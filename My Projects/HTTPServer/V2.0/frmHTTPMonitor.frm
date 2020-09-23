VERSION 5.00
Begin VB.Form frmHTTPMonitor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "HTTP Monitor"
   ClientHeight    =   6720
   ClientLeft      =   7995
   ClientTop       =   1680
   ClientWidth     =   6570
   Icon            =   "frmHTTPMonitor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtDebug 
      BorderStyle     =   0  'None
      Height          =   6285
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   390
      Width           =   6525
   End
   Begin VB.Label lblConnections 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Connections: none"
      Height          =   195
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   1335
   End
End
Attribute VB_Name = "frmHTTPMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub WriteStatus(Status As String)

  If txtDebug.Text <> "" Then
    txtDebug.Text = txtDebug.Text & vbNewLine & Status
  Else
    txtDebug.Text = txtDebug.Text & Status
  End If

End Sub

Private Sub Form_Resize()

  txtDebug.Move 0, lblConnections.Top + lblConnections.Height, ScaleWidth, ScaleHeight - (lblConnections.Top + lblConnections.Height)

End Sub


Private Sub txtDebug_DblClick()

  txtDebug.Text = ""

End Sub


