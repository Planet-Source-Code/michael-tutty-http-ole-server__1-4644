VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HTTPHeaderList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"HTTPParameter"
Attribute VB_Ext_KEY = "Member0" ,"HTTPParameter"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mCol As Collection
Public Function AddItems(Request As String, Optional NewCol As Boolean = False) As Boolean

  Dim col As Collection
  Dim objNew As HTTPHeader
  Dim iLoop As Long
  
  On Error GoTo AddItemsError
  
  If NewCol Then
    Set mCol = New Collection
  End If
  Set col = HeaderDecode(Request)
  
  For iLoop = 1 To col.Count
    Set objNew = New HTTPHeader
    If objNew.Initialize(col(iLoop)) = False Then
      'Debug.Assert False
      Exit Function
    End If
    mCol.Add objNew
    Set objNew = Nothing
  Next iLoop
  
  AddItems = True
  
AddItemsError:
  Debug.Assert Err.Number = 0

End Function

Public Property Get Item(vntIndexKey As Variant) As HTTPHeader
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Public Function Values(Key As String) As Variant

  Dim col As Collection
  Dim RetVal As String
  Dim CurVal As String
  Dim iCount As Long
  Dim iLoop As Long
  
  For iLoop = 1 To mCol.Count
    If StrComp(mCol(iLoop).Name, Key, vbTextCompare) = 0 Then
      CurVal = mCol(iLoop).Value
      iCount = iCount + 1
      Select Case iCount
        Case 1
          RetVal = CurVal
          
        Case 2
          Set col = New Collection
          col.Add RetVal
          col.Add CurVal
          
        Case Else
          col.Add CurVal
      End Select
    End If
  Next iLoop
  
  If iCount > 1 Then
    Set Values = col
    Set col = Nothing
  Else
    Values = RetVal
  End If

End Function
Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub
Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

