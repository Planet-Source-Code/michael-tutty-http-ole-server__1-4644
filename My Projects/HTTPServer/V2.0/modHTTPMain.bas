Attribute VB_Name = "modHTTPServMain"
Option Explicit

Public Const BIND_PORT_DEFAULT = 80

Private iConnections As Long

Public Sub RemoveConnection()

  iConnections = iConnections - 1

End Sub

Public Sub AddConnection()

  iConnections = iConnections + 1

End Sub

Public Sub Main()

  ' Do Nothing

End Sub


Public Sub WriteStatus(Status As String)

End Sub
