VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "People"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_People As Collection

Public Sub Add(ByRef p As Person)
    m_People.Add p
End Sub


Public Sub Update(ByVal index As Integer, ByRef updated As Person)
    Dim existing As Person
    
    Set existing = m_People(index)
    With existing
        .Name = updated.Name
        .Surname = updated.Surname
    End With
End Sub


Public Sub Delete(ByVal index As Integer)
    m_People.Remove index
End Sub


Public Function GetPersonByIndex(ByVal index As Integer) As Person
    Set GetPersonByIndex = m_People(index)
End Function


Public Function Count() As Integer
    Count = m_People.Count
End Function


Public Function GetIterator(Optional ByVal filter As IPersonFilter = Nothing) As PersonIterator
    Dim result As New PersonIterator
    Set result.Source = Me
    Set result.filter = filter
    Set GetIterator = result
End Function

Private Sub Class_Initialize()
    Set m_People = New Collection
End Sub

