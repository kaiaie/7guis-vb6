VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Queue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Queue As New Collection

Public Sub Enqueue(ByVal Value As Variant)
    If m_Queue.Count = 0 Then
        m_Queue.Add Value
    Else
        m_Queue.Add Value, Before:=1
    End If
End Sub


Public Function Dequeue() As Variant
    If Count = 0 Then
        Err.Raise 5, _
            Source:="Queue::Dequeue", _
            Description:="Queue is empty"
    End If
    Dequeue = m_Queue(Count)
    m_Queue.Remove Count
End Function


Public Property Get Count() As Long
    Count = m_Queue.Count
End Property


Public Sub Clear()
    Do While Count > 0
        m_Queue.Remove 0
    Loop
End Sub


Public Function Clone() As Queue
    Dim result As New Queue
    Dim temp As New Stack
    Dim v As Variant
    For Each v In m_Queue
        temp.Push v
    Next
    Do Until temp.IsEmpty
        result.Enqueue temp.Pop()
    Loop
    Set Clone = result
End Function
