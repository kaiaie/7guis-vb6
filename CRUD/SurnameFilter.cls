VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SurnameFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IPersonFilter

Private m_SurnamePrefix As String


Public Property Get SurnamePrefix() As String
    SurnamePrefix = m_SurnamePrefix
End Property


Public Property Let SurnamePrefix(ByVal value As String)
    m_SurnamePrefix = value
End Property


Private Function IPersonFilter_ApplyFilter(p As Person) As Boolean
    If Len(SurnamePrefix) = 0 Then
        IPersonFilter_ApplyFilter = True
    ElseIf Len(SurnamePrefix) > Len(p.Surname) Then
        IPersonFilter_ApplyFilter = False
    Else
        IPersonFilter_ApplyFilter = UCase(Left(p.Surname, Len(SurnamePrefix))) = UCase(SurnamePrefix)
    End If
End Function
