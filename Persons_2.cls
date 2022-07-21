VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Persons_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Items_ As Collection

Private Sub Class_Initialize()
    Set Items = New Collection
End Sub

Public Sub Add(ByVal newId As String, ByVal newName As String, ByVal newAge As Long)
    Dim p As Person: Set p = New Person
    With p
        .Id = newId
        .FirstName = newName
        .Age = newAge
    End With
    
    Items.Add p, newId
End Sub
