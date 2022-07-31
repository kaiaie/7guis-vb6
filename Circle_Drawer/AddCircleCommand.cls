VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AddCircleCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ICommand

Private m_TargetForm As DrawCircleForm
Private m_CircleIndex As Integer
Private m_CircleX As Single
Private m_CircleY As Single

Private Sub ICommand_Forward()
    Forward
End Sub

Private Sub ICommand_Reverse()
    Reverse
End Sub

Private Property Set ICommand_TargetForm(RHS As DrawCircleForm)
    Set m_TargetForm = RHS
End Property


Public Property Set TargetForm(RHS As DrawCircleForm)
    Set m_TargetForm = RHS
End Property


Public Property Get TargetForm() As DrawCircleForm
    Set TargetForm = m_TargetForm
End Property


Public Property Get CircleIndex() As Integer
    CircleIndex = m_CircleIndex
End Property


Public Property Let CircleIndex(ByVal Value As Integer)
    m_CircleIndex = Value
End Property


Public Property Get CircleX() As Single
    CircleX = m_CircleX
End Property


Public Property Let CircleX(ByVal Value As Single)
    m_CircleX = Value
End Property


Public Property Get CircleY() As Single
    CircleY = m_CircleY
End Property


Public Property Let CircleY(ByRef Value As Single)
    m_CircleY = Value
End Property


Public Sub Forward()
    TargetForm.AddCircle CircleIndex, CircleX, CircleY
End Sub


Public Sub Reverse()
    TargetForm.RemoveCircle CircleIndex
End Sub
