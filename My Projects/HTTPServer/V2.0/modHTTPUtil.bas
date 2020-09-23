Attribute VB_Name = "modHTTPUtil"
Option Explicit

Public Function GetDecodedChar(ByVal sEncChar As String) As Integer

  Dim nChar As Integer
  
  If Left(sEncChar, 1) <> "%" Then
    GetDecodedChar = sEncChar
    Exit Function
  End If
  
  GetDecodedChar = CLng("&H" & Mid(sEncChar, 2))

End Function

Public Function GetEncodedChar(nChar As Integer) As String

  Dim sTemp As String

  If nChar < 0 Then
    Exit Function
  End If
  
  If nChar > 255 Then
    Exit Function
  End If
  
  sTemp = Hex(nChar)
  
  If Len(sTemp) < 2 Then
    sTemp = "0" & sTemp
  End If
  
  GetEncodedChar = "%" & sTemp

End Function
Public Function HeaderDecode(ByVal Data As String) As Collection

  Dim col As New Collection
  Dim iChar As Long
  Dim HeaderItem  As String

  iChar = 1

  Do While Len(Data)
    HeaderItem = SplitFrom(Data, vbCr)
    If Len(HeaderItem) = 0 Then
      HeaderItem = Data
      Data = ""
    End If
    col.Add HeaderItem
  Loop
  
  Set HeaderDecode = col
  Set col = Nothing

End Function

Public Function SplitFrom(StringToSplit, Delimiter) As String

  Dim iFind As Long

  iFind = InStr(1, StringToSplit, Delimiter, vbTextCompare)
  
  If iFind = 0 Then
    Exit Function
  End If
  
  SplitFrom = Left(StringToSplit, iFind - 1)
  StringToSplit = Mid(StringToSplit, iFind + Len(Delimiter))

End Function

Public Function URLDecode(ByVal Data As String) As Collection

  Dim col As New Collection
  Dim iChar As Long
  Dim nChar As String

  iChar = 1

  Do While iChar < Len(Data)
    nChar = Mid(Data, iChar, 1)
    Select Case nChar
      Case "%"
        nChar = Mid(Data, iChar, 3)
        Data = Left(Data, iChar - 1) & Chr(GetDecodedChar(nChar)) & Mid(Data, iChar + Len(nChar))
        iChar = iChar + 2
      
      Case "+"
        Data = Left(Data, iChar - 1) & " " & Mid(Data, iChar + 1)
        iChar = iChar + 1

      Case "&"
        col.Add Left(Data, iChar - 1)
        Data = Mid(Data, iChar + 1)
        iChar = 1
      
      Case Else
        iChar = iChar + 1
    
    End Select
  Loop
  If Len(Data) Then
    col.Add Data
  End If

  Set URLDecode = col
  Set col = Nothing

End Function
Public Sub DebugPrint(Request As HTTPServiceRequest, Optional Data As String = "")

  Dim sTemp As String
  Dim iLoop As Long
  
  sTemp = "Debug Dump for HTTPServiceRequest, Index " & Request.Index & ": " & vbNewLine
  
    sTemp = sTemp & Space(4) & "HTTP Version: " & Request.HTTPVersion & vbNewLine
    sTemp = sTemp & Space(4) & "Request Method: " & Request.Method & vbNewLine
    sTemp = sTemp & Space(4) & "Requested URI: " & Request.URI & vbNewLine
    sTemp = sTemp & Space(4) & "Header List: " & vbNewLine
    With Request.Headers
      For iLoop = 1 To .Count
        sTemp = sTemp & Space(8) & "Name: " & .Item(iLoop).Name & ", Value: " & .Item(iLoop).Value & vbNewLine
      Next iLoop
    End With
    
    sTemp = sTemp & Space(4) & "Parameter List: " & vbNewLine
    With Request.Parameters
      For iLoop = 1 To .Count
        sTemp = sTemp & Space(8) & "Name: " & .Item(iLoop).Name & ", Value: " & .Item(iLoop).Value & vbNewLine
      Next iLoop
    End With
    
    If Data <> "" Then
      sTemp = sTemp & vbNewLine & "Raw HTTP:" & vbNewLine & Data & vbNewLine
    End If
    
    WriteStatus sTemp

End Sub


Public Function GetHTTPOutput(Content As String, _
                                                                    Optional Headers As HTTPParameterList, _
                                                                    Optional HTTPVersion As String = "HTTP/1.0", _
                                                                    Optional Code As String = "200", _
                                                                    Optional Desc As String = "Document follows", _
                                                                    Optional ContentType As String = "text/html", _
                                                                    Optional ServerType As String = "WTServer") _
                                                                    As String
                                                                    
  Dim OutString As New DelimitedString
  Dim iLoop As Long
  
  With OutString
    .Delimiter = vbNewLine
    .Add HTTPVersion & " " & Code & " " & Desc
    .Add "Content-type: " & ContentType
    .Add "Content-length: " & Len(Content)
    .Add "Server: " & ServerType
    If Not (Headers Is Nothing) Then
      For iLoop = 1 To Headers.Count
        .Add Headers(iLoop).Name & ": " & Headers(iLoop).Value
      Next iLoop
    End If
    .Add ""
    .Add Content
  End With

  GetHTTPOutput = OutString.Value & Chr(0)

  Set OutString = Nothing

End Function
Public Function URLEncode(ByVal Data As String) As String

  Dim iChar As Long
  Dim nChar As Integer
  
  iChar = 1

  Do While iChar < Len(Data)
    nChar = Asc(Mid(Data, iChar, 1))
    Select Case nChar
      ' Case 0 To 47
        ' Use Case Else
      
      Case 48 To 57   ' 0 to 9
        ' Leave alone
      
      ' Case 58 To 64
        ' Use Case Else
      
      Case 65 To 90   ' A-Z (uppercase)
        ' Leave alone
      
      ' Case 91 To 96
        ' Use Case Else
      
      Case 97 To 122    ' a-z (lowercase)
        ' Leave alone
      
      Case Else
        Data = Left(Data, iChar - 1) & GetEncodedChar(nChar) & Mid(Data, iChar + 1)
        iChar = iChar + 2
      
    End Select
  Loop
  
  URLEncode = Data

End Function
