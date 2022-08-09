VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CellInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum CellContents
    Literal
    Formula
End Enum

Private m_CellType As CellContents
Private m_Value As Variant
Private m_Expression As Expression
Private m_Model As GridModel


Public Property Get CellType() As CellContents
    CellType = m_CellType
End Property


Public Property Let CellType(ByVal Value As CellContents)
    m_CellType = Value
End Property


Public Property Get CellValue() As Variant
    CellValue = m_Value
End Property


Public Property Let CellValue(ByVal Value As Variant)
    m_Value = Value
    ' TODO: Start recalc
End Property


Public Property Get CellFormula() As String
    If CellType <> Formula Then
        Err.Raise 5, Description:="The cell does not contain a formula"
    End If
    CellFormula = m_Expression.Expression
End Property


Public Property Let CellFormula(ByVal Value As String)
    Dim r As New Resolver
    Set r.Model = Model
    Set m_Expression = New Expression
    Set m_Expression.Resolver = r
    m_Expression.Expression = Value
    CellValue = m_Expression.Evaluate()
End Property


Public Property Get Model() As GridModel
    Set Model = m_Model
End Property


Public Property Set Model(ByRef Value As GridModel)
    Set m_Model = Value
End Property


Private Sub Class_Terminate()
    Set m_Model = Nothing
End Sub


Public Property Get Dependencies() As Collection
    If CellType = Formula Then
        Set Dependencies = m_Expression.ExternalReferences
    End If
End Property


Public Sub Recalculate()
    If CellType <> Formula Then Exit Sub
    CellValue = m_Expression.Evaluate()
End Sub
