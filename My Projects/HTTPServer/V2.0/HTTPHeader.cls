VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HTTPHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mstrName As String
Private mstrValue As String
Friend Function Initialize(ByVal InputLine As String) As Boolean

  mstrName = Trim(SplitFrom(InputLine, ":"))
  mstrValue = Trim(InputLine)
  Initialize = (Len(mstrName) > 0)

End Function

Public Property Get Name() As String

  Name = mstrName

End Property


Public Property Get Value() As String

  Value = mstrValue

End Property
