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
''' \brief A simple graph implementation
Option Explicit

Private Const FROM_NODE As Integer = 0
Private Const TO_NODE As Integer = 1

Private m_NodeCount As Integer
Private m_Nodes() As String
Private m_EdgeCount As Integer
Private m_Edges() As Variant


Public Sub AddNode(ByVal NodeName As String)
    NodeName = Trim(UCase(NodeName))
    If Len(NodeName) = 0 Then
        Err.Raise 5, _
            Source:="Graph::AddNode", _
            Description:="The node name must not be empty"
    End If
    ' Note: debatable if trying to add a duplicate node or edge to the graph
    ' is an error or not. Decided to go trying to add a duplicate node is an
    ' error but an edge is not, but the case could be made for both or
    ' neither...
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
    FromNode = Trim(UCase(FromNode))
    ToNode = Trim(UCase(ToNode))
    If Len(FromNode) = 0 Then
        Err.Raise 5, _
            Source:="Graph::AddEdge", _
            Description:="The ""from"" node name must not be empty"
    End If
    If Len(ToNode) = 0 Then
        Err.Raise 5, _
            Source:="Graph::AddEdge", _
            Description:="The ""to"" node name must not be empty"
    End If
    If Not NodeExists(FromNode) Then
        AddNode FromNode
    End If
    If Not NodeExists(ToNode) Then
        AddNode ToNode
    End If
    If EdgeExists(FromNode, ToNode) Then
        Exit Sub
    End If
    m_EdgeCount = m_EdgeCount + 1
    ReDim Preserve m_Edges(1 To m_EdgeCount)
    m_Edges(m_EdgeCount) = Array(FromNode, ToNode)
End Sub


''' \brief Returns the graph's nodes in topologically sorted order
''' \remarks Uses the Kahn algorithm described in:
''' https://en.wikipedia.org/wiki/Topological_sorting#Kahn%27s_algorithm
Public Function TopologicalSort() As Collection
    Dim result As New Collection
    If m_NodeCount = 0 Then
        Set TopologicalSort = result
        Exit Function
    End If
    Dim edges As New Collection
    Dim i As Integer, e As Variant, v As Variant
    
    ' Make a copy of the list of edges (entries in this list are
    ' deleted as the algorithm progesses)
    For i = LBound(m_Edges) To UBound(m_Edges)
        e = m_Edges(i)
        edges.Add e
    Next
    ' Get all the nodes that have no incoming edges
    Dim startNodes As New Queue
    For i = LBound(m_Nodes) To UBound(m_Nodes)
        If EdgesTo(m_Nodes(i)) = 0 Then
            startNodes.Enqueue m_Nodes(i)
        End If
    Next
    Do While startNodes.Count > 0
        Dim n As String
        n = CStr(startNodes.Dequeue())
        result.Add n
        ' Any edges coming from this node? If so, get the list of nodes it's
        ' connected to
        Dim ms As New Collection
        For Each e In edges
            If CStr(e(FROM_NODE)) = n Then
                ms.Add e(TO_NODE)
            End If
        Next
        ' For each of these nodes...
        For Each v In ms
            ' Any edges from the start node to this? If so, remove it
            For i = 1 To edges.Count
                e = edges(i)
                If CStr(e(FROM_NODE)) = n And CStr(e(TO_NODE)) = CStr(v) Then
                    edges.Remove i
                    Exit For
                End If
            Next
            ' If there are no more edges to this node now, add it to the queue
            ' of starting nodes
            If EdgesTo(CStr(v)) = 0 Then
                startNodes.Enqueue v
            End If
        Next
    Loop
    If edges.Count > 0 Then
        Err.Raise vbObjectError + 990, _
            Source:="Graph::TopologicalSort", _
            Description:="Cycle detected in graph"
    End If
    Set TopologicalSort = result
End Function


Public Property Get NodeCount() As Integer
    NodeCount = m_NodeCount
End Property


Public Property Get EdgeCount() As Integer
    EdgeCount = m_EdgeCount
End Property


Public Function NodeExists(ByVal NodeName As String) As Boolean
    NodeName = Trim(UCase(NodeName))
    If Len(NodeName) = 0 Then
        Err.Raise 5, _
            Source:="Graph::NodeExists", _
            Description:="The node name must not be empty"
    End If
    If m_NodeCount = 0 Then
        NodeExists = False
        Exit Function
    End If
    Dim i As Integer
    For i = LBound(m_Nodes) To UBound(m_Nodes)
        If NodeName = m_Nodes(i) Then
            NodeExists = True
            Exit Function
        End If
    Next
End Function


Public Function EdgeExists(ByVal FromNode As String, ToNode As String) As Boolean
    FromNode = Trim(UCase(FromNode))
    ToNode = Trim(UCase(ToNode))
    If Len(FromNode) = 0 Then
        Err.Raise 5, _
            Source:="Graph::EdgeExists", _
            Description:="The ""from"" node name must not be empty"
    End If
    If Len(ToNode) = 0 Then
        Err.Raise 5, _
            Source:="Graph::EdgeExists", _
            Description:="The ""to"" node name must not be empty"
    End If
    If m_EdgeCount = 0 Then
        EdgeExists = False
        Exit Function
    End If
    Dim i As Integer, e As Variant
    For i = LBound(m_Edges) To UBound(m_Edges)
        e = m_Edges(i)
        If e(FROM_NODE) = FromNode And e(TO_NODE) = ToNode Then
            EdgeExists = True
            Exit Function
        End If
    Next
End Function


''' \brief Returns the number of incoming edges to the specified node
Public Function EdgesTo(ByVal NodeName As String) As Integer
    NodeName = Trim(UCase(NodeName))
    If Len(NodeName) = 0 Then
        Err.Raise 5, _
            Source:="Graph::EdgesTo", _
            Description:="The node name must not be empty"
    End If
    Dim i As Integer, e As Variant
    For i = LBound(m_Edges) To UBound(m_Edges)
        e = m_Edges(i)
        If e(TO_NODE) = NodeName Then
            EdgesTo = EdgesTo + 1
        End If
    Next
End Function

