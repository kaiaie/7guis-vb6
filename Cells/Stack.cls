VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Stack As New Collection

Public Sub Push(ByVal Value As Variant)
    m_Stack.Add Value
End Sub


Public Function Pop() As Variant
    If Count = 0 Then
        Err.Raise 5, _
            Source:="Stack::Pop", _
            Description:="Stack is empty"
    End If
    Pop = m_Stack(Count)
    m_Stack.Remove Count
End Function


Public Property Get Count() As Long
    Count = m_Stack.Count
End Property


Public Sub Clear()
    Do While m_Stack.Count > 0
        m_Stack.Remove 0
    Loop
End Sub


Public Function Peek() As Variant
    If Count = 0 Then
        Err.Raise 5, _
            Source:="Stack::Pop", _
            Description:="Stack is empty"
    End If
    Peek = m_Stack(Count)
End Function
