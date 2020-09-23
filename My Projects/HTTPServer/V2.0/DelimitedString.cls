VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DelimitedString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Delimiter As String
Private mstrString As String

Private Const DEFAULT_DELIMITER As String = vbTab

Public Sub Add(NewItem As String)

  If mstrString <> "" Then
    mstrString = mstrString & Delimiter & NewItem
  Else
    mstrString = mstrString & NewItem
  End If

End Sub

Public Sub Clear()

  mstrString = ""

End Sub

Public Property Get Value() As String

  Value = mstrString

End Property


Private Sub Class_Initialize()

  Delimiter = DEFAULT_DELIMITER

End Sub


