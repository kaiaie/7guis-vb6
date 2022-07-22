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

Private m_currentIdx As Integer
Private m_Source As People
Private m_Filter As IPersonFilter

Public Property Set Source(ByRef value As People)
    Set m_Source = value
    m_currentIdx = 1
End Property

Public Property Get Filter() As IPersonFilter
    Set Filter = m_Filter
End Property

Public Property Set Filter(ByRef value As IPersonFilter)
    Set m_Filter = value
End Property

Public Property Get Current() As Person
    Set Current = m_Source.GetPersonByIndex(m_currentIdx)
End Property

Public Sub MoveNext()
    Do
        m_currentIdx = m_currentIdx + 1
        If m_currentIdx > m_Source.Count Then
            Exit Sub
        End If
        If Filter Is Nothing Then
            Exit Sub
        End If
        If Filter.ApplyFilter(Current) Then
            Exit Sub
        End If
    Loop
End Sub

Public Property Get EOF() As Boolean
    EOF = m_currentIdx > m_Source.Count
End Property
