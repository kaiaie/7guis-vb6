VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PersonIterator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'* \brief Allows the list of people to be traversed, with an optional filter.
'* Modelled somewhat on the ADO Recordset class, but for brevity, only forward
'* motion through the list is supported
'*
'* \remarks Really missing .NET iterators here! It can be done by low-level
'* COM hackery (http://www.vb-helper.com/howto_custom_collection_with_for_each.html)
'* but will do for now
'* \note Needs to be cleaned up but too much time wasted on it

Private m_CurrentIndex As Integer
Private m_Source As People
Private m_Filter As IPersonFilter
Private m_BOF As Boolean
Private m_Current As Person
Private m_Next As Person

'* The repository being traversed
Public Property Set Source(ByRef value As People)
    Set m_Source = value
    m_BOF = True
End Property


'* The filter that is applied to each record as it is fetched
Public Property Get Filter() As IPersonFilter
    Set Filter = m_Filter
End Property


Public Property Set Filter(ByRef value As IPersonFilter)
    If Not m_Filter Is Nothing Then
        Err.Raise vbObjectError + 900, _
            Description:="The filter cannot be changed; instead, create a new iterator with the new filter"
    End If
    Set m_Filter = value
End Property


Public Property Get Current() As Person
    Set Current = m_Current
End Property

'* Starts the iteration
Public Sub MoveFirst()
    Dim filterResult As Boolean
    
    m_BOF = True
    Set m_Current = Nothing
    Set m_Next = Nothing
    
    ' Check if source is empty
    If m_Source.Count = 0 Then
        Exit Sub
    End If
    
    m_CurrentIndex = 1
    Dim candidate As Person
getCurr:
    Set candidate = m_Source.GetPersonByIndex(m_CurrentIndex)
    ' Check if record satisfies filter
    If Not Filter Is Nothing Then
        filterResult = Filter.ApplyFilter(candidate)
        If Not filterResult Then
            Set candidate = Nothing
        End If
    End If
    If Not candidate Is Nothing Then
        Set m_Current = candidate
        m_CurrentIndex = m_CurrentIndex + 1
        If m_CurrentIndex > m_Source.Count Then Exit Sub
        GoTo getNext
    Else
        m_CurrentIndex = m_CurrentIndex + 1
        If m_CurrentIndex > m_Source.Count Then Exit Sub
        GoTo getCurr
    End If
getNext:
    Set candidate = m_Source.GetPersonByIndex(m_CurrentIndex)
    ' Check if record satisfies filter
    If Not Filter Is Nothing Then
        filterResult = Filter.ApplyFilter(candidate)
        If Not filterResult Then
            Set candidate = Nothing
        End If
    End If
    If Not candidate Is Nothing Then
        Set m_Next = candidate
    Else
        m_CurrentIndex = m_CurrentIndex + 1
        If m_CurrentIndex > m_Source.Count Then Exit Sub
        GoTo getNext
    End If
End Sub

'* Moves the Current property to the next record in the underlying repository
'* that meets the filter criteria
Public Sub MoveNext()
    Set m_Current = m_Next
    Set m_Next = Nothing
    
    m_CurrentIndex = m_CurrentIndex + 1
    If m_CurrentIndex > m_Source.Count Then Exit Sub
    
    Dim candidate As Person
getNext:
    Set candidate = m_Source.GetPersonByIndex(m_CurrentIndex)
    ' Check if record satisfies filter
    If Not Filter Is Nothing Then
        Dim filterResult As Boolean
        filterResult = Filter.ApplyFilter(candidate)
        If Not filterResult Then
            Set candidate = Nothing
        End If
    End If
    If Not candidate Is Nothing Then
        Set m_Next = candidate
    Else
        m_CurrentIndex = m_CurrentIndex + 1
        If m_CurrentIndex > m_Source.Count Then Exit Sub
        GoTo getNext
    End If
End Sub


'* Returns true when at the start of the list
Public Property Get BOF() As Boolean
    BOF = m_BOF
End Property


'* Returns true when the end of the list has been reached
Public Property Get EOF() As Boolean
    EOF = (m_Current Is Nothing)
End Property
