VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Graph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type EdgeType
    FromNode As String
    ToNode As String
End Type

Private m_NodeCount As Integer
Private m_Nodes() As String
Private m_EdgeCount As Integer
Private m_Edges() As EdgeType


Public Sub AddNode(ByVal NodeName As String)
    If NodeExists(NodeName) Then
        Err.Raise 5, _
            Source:="Graph::AddNode", _
            Description:=Strings.Format("A node with the name ""{0}"" already exists in the graph", NodeName)
    End If
    m_NodeCount = m_NodeCount + 1
    ReDim Preserve m_Nodes(1 To m_NodeCount)
    m_Nodes(m_NodeCount) = NodeName
End Sub


Public Sub AddEdge(ByVal FromNode As String, ToNode As String)
    Dim t As EdgeType
    t.FromNode = FromNode
    t.ToNode = ToNode
    m_EdgeCount = m_EdgeCount + 1
    ReDim Preserve m_Edges(1 To m_EdgeCount)
    m_Edges(m_EdgeCount) = t
End Sub


Public Function TopologicalSort() As Collection
    Dim result As New Collection
End Function


Public Property Get NodeCount() As Integer
    NodeCount = m_NodeCount
End Property


Public Property Get EdgeCount() As Integer
    EdgeCount = m_EdgeCount
End Property


Private Function NodeExists(ByVal NodeName As String) As Boolean
    Dim i As Integer
    For i = LBound(m_Nodes) To UBound(m_Nodes)
        If UCase(NodeName) = UCase(m_Nodes(i)) Then
            NodeExists = True
            Exit Function
        End If
    Next
End Function
