VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CollStrC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'Option Explicit

' Copyright ©2003 by mrk
' Keinerlei Gewährleistung für Software und Dokumentation. Die Nutzung erfolgt uneingeschränkt auf eigenes Risiko. Der Autor kann für jedweden auftretenden Schaden oder Folgeschaden nicht haftbar gemacht werden.

Public Enum estrStatus
  OK = -1           'zustimmen
  NameToShort = 1   'Name zu kurz
  NameExists = 2    'Name existiert
  NameUnlocated = 3 'Name nicht gefunden
End Enum

Private c As Collection
Private propStatus As estrStatus

Private Sub Class_Initialize()
  Set c = New Collection
End Sub
Private Sub Class_Terminate()
Clear
Set c = Nothing
End Sub

Public Property Get Status()
  Status = propStatus
End Property

Public Function Add( _
             ByVal Name As String, _
             ByVal Item As String) As estrStatus
 If Len(Name) = 0 Then
    propStatus = estrStatus.NameToShort
Exit Function
 End If
 
  If Exist(Name) Then
    propStatus = estrStatus.NameExists
    Exit Function
 End If
 
  c.Add Item, Name
 propStatus = estrStatus.OK
 Add = estrStatus.OK
End Function

Public Property Get Item( _
              ByVal Name As String) As String
  If Len(Name) = 0& Then
     Item = ""
     propStatus = estrStatus.NameToShort
Exit Property
  End If

  If Not Exist(Name) Then
     Item = ""
     propStatus = estrStatus.NameUnlocated
Exit Property
  End If

  Item = c.Item(Name)
End Property

Public Property Let Item( _
              ByVal Name As String, _
              ByVal Item As String)
  If Len(Name) = 0& Then
     propStatus = estrStatus.NameToShort
Exit Property
  End If

  If Not Exist(Name) Then
     propStatus = estrStatus.NameUnlocated
Exit Property
  End If

  c.Remove Name
  c.Add Item, Name

  propStatus = estrStatus.OK
End Property

Public Function Exist(ByRef Name As String) As Boolean
  On Error Resume Next
  
  Dim t As String
  
  t = c.Item(Name)
  If CBool(Err.Number = 0) Then
     propStatus = estrStatus.OK
     Exist = True
    
    Else
     propStatus = estrStatus.NameUnlocated
     Exist = False
     Err.Clear
  End If
  On Error GoTo 0
End Function

Public Function Count() As Long
  Count = c.Count
  propStatus = estrStatus.OK
End Function

Public Sub Clear()
  Dim i As Long
  Dim u As Long
  
  With c
    u = .Count
    For i = 1 To u
        .Remove 1&
    Next i
  End With
  propStatus = estrStatus.OK
  End Sub

