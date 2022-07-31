VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SetDiameterCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ICommand

Private m_TargetForm As DrawCircleForm
Private m_CircleIndex As Integer
Private m_OldDiameter As Integer
Private m_NewDiameter As Integer

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


Private Property Get TargetForm() As DrawCircleForm
    Set TargetForm = m_TargetForm
End Property


Public Property Get CircleIndex() As Integer
    CircleIndex = m_CircleIndex
End Property


Public Property Let CircleIndex(ByVal Value As Integer)
    m_CircleIndex = Value
End Property


Public Property Get OldDiameter() As Single
    OldDiameter = m_OldDiameter
End Property


Public Property Let OldDiameter(ByVal Value As Single)
    m_OldDiameter = Value
End Property


Public Property Get NewDiameter() As Single
    NewDiameter = m_NewDiameter
End Property


Public Property Let NewDiameter(ByVal Value As Single)
    m_NewDiameter = Value
End Property


Public Sub Forward()
    TargetForm.AdjustDiameter CircleIndex, NewDiameter
End Sub


Public Sub Reverse()
    TargetForm.AdjustDiameter CircleIndex, OldDiameter
End Sub

