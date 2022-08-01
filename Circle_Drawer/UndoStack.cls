VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UndoStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim m_Actions() As ICommand
Dim m_ActionCount As Integer
Dim m_CurrentIndex As Integer

Private m_Form As DrawCircleForm


Public Property Get OwnerForm() As DrawCircleForm
    Set OwnerForm = m_Form
End Property


Public Property Set OwnerForm(ByRef Value As DrawCircleForm)
    Set m_Form = Value
End Property


Public Property Get CanUndo() As Boolean
    CanUndo = m_CurrentIndex > 0
End Property


Public Property Get CanRedo() As Boolean
    CanRedo = m_CurrentIndex <> m_ActionCount
End Property


Public Sub Undo()
    m_Actions(m_CurrentIndex).Reverse
    m_CurrentIndex = m_CurrentIndex - 1
    Debug.Print Strings.Format("UndoStack::Undo: {{ count: {0}, current: {1}}", m_ActionCount, m_CurrentIndex)
End Sub


Public Sub Redo()
    m_Actions(m_CurrentIndex + 1).Forward
    m_CurrentIndex = m_CurrentIndex + 1
    Debug.Print Strings.Format("UndoStack::Redo: {{ count: {0}, current: {1}}", m_ActionCount, m_CurrentIndex)
End Sub


Private Sub Class_Terminate()
    Set m_Form = Nothing
End Sub


Public Sub AddDiameterAdjustment(ByVal Index As Integer, ByVal OldDiameter As Single, ByVal NewDiameter As Single)
    Dim cmd As New SetDiameterCommand
    With cmd
        Set .TargetForm = Me.OwnerForm
        .CircleIndex = Index
        .OldDiameter = OldDiameter
        .NewDiameter = NewDiameter
    End With
    Add cmd
End Sub


Public Sub AddCircle(ByVal Index As Integer, ByVal X As Single, ByVal Y As Single)
    Dim cmd As New AddCircleCommand
    With cmd
        Set .TargetForm = Me.OwnerForm
        .CircleIndex = Index
        .CircleX = X
        .CircleY = Y
    End With
    Debug.Print Strings.Format("UndoStack::AddCircle: {{ index: {0}, x: {1}, y:{2} }", Index, X, Y)
    Add cmd
End Sub


Private Sub Add(ByRef cmd As ICommand)
    m_CurrentIndex = m_CurrentIndex + 1
    m_ActionCount = m_CurrentIndex
    ReDim Preserve m_Actions(1 To m_CurrentIndex)
    Set m_Actions(m_CurrentIndex) = cmd
    Debug.Print Strings.Format("UndoStack::Add: {{ count: {0}, current: {1}}", m_ActionCount, m_CurrentIndex)
End Sub
