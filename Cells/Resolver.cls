VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Resolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
''' \brief Resolves cell references

Private m_Model As GridModel

Public Property Get Model() As GridModel
    Set Model = m_Model
End Property


Public Property Set Model(ByRef Value As GridModel)
    Set m_Model = Value
End Property


Public Function Resolve(ByVal cellAddress As String) As Variant
    Resolve = Model.GetCellInfo(cellAddress).CellValue
End Function

Private Sub Class_Terminate()
    Set Model = Nothing
End Sub
