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
'* A repository for Person objects
'*
'* \remarks This is not as clean as I'd like, owing to shortcomings with
'* COM Collection objects. It would be nicer to use either a real in-memory
'* database or a data structure like a tree, but it would take too long!

Private m_Id As Integer
Private m_People As Collection

'* \brief Adds a person to the repository
Public Sub Add(ByRef p As Person)
    If p Is Nothing Then
        Err.Raise vbObjectError + 1901, _
            Description:="An empty person cannot be added"
    End If
    p.id = m_Id
    m_People.Add p, ToKey(m_Id)
    m_Id = m_Id + 1
End Sub


'* \brief Updates the details of an existing person
Public Sub Update(ByRef updated As Person)
    Dim existing As Person
    
    For Each existing In m_People
        If existing.id = updated.id Then
            existing.Name = updated.Name
            existing.Surname = updated.Surname
            Exit Sub
        End If
    Next
    Err.Raise vbObjectError + 1900, _
        Description:="Person not found"
End Sub


'* Deletes a person from the repository
Public Sub Delete(ByVal id As Integer)
    Dim i As Integer
    Dim p As Person
    For i = 1 To m_People.Count
        Set p = m_People(i)
        If p.id = id Then
            m_People.Remove i
            Exit Sub
        End If
    Next
End Sub


'* \brief Returns the nth person in the repository by order of insertion
Public Function GetPersonByIndex(ByVal index As Integer) As Person
    Set GetPersonByIndex = m_People(index)
End Function


'* \brief Returns the person with the specified ID
'* IDs are assigned in ascending order by insertion but remain the
'* same regardless of deletions, unlike GetPersonByIndex
Public Function GetPersonByID(ByVal id As Integer) As Person
    Dim key As String
    Dim result As Person
    key = ToKey(id)
    On Error Resume Next
    Set result = m_People(key)
    On Error GoTo 0
    If Not result Is Nothing Then
        Set GetPersonByID = result
        Exit Function
    End If
    Err.Raise vbObjectError + 1900, _
        Description:="Person not found"
End Function


'* Returns the number of people in the repository
Public Function Count() As Integer
    Count = m_People.Count
End Function


'* Returns an object that allows the list of people to be traversed
Public Function GetIterator(Optional ByVal Filter As IPersonFilter = Nothing) As PersonIterator
    Dim result As New PersonIterator
    Set result.Source = Me
    Set result.Filter = Filter
    result.MoveFirst
    Set GetIterator = result
End Function


Private Sub Class_Initialize()
    Set m_People = New Collection
End Sub


Private Function ToKey(ByVal id As Integer)
    ToKey = "P" & Format(id, String(5, "0"))
End Function


Private Sub Class_Terminate()
    Set m_People = Nothing
End Sub
