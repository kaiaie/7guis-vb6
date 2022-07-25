VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Person"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'* A Plain Old COM Object to represent the personal details that can be
'* edited in the user interface

Private m_Id As Integer
Private m_Name As String
Private m_Surname As String

Public Property Get Name() As String
    Name = m_Name
End Property


Public Property Let Name(ByVal value As String)
    m_Name = value
End Property


Public Property Get Surname() As String
    Surname = m_Surname
End Property


Public Property Let Surname(ByVal value As String)
    m_Surname = value
End Property


Public Property Get id() As Integer
    id = m_Id
End Property


Public Property Let id(ByVal value As Integer)
    m_Id = value
End Property
