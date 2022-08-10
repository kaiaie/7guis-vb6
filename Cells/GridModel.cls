VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GridModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Rows As New Collection
Private m_Cols As New Collection
Private m_Cells As New Collection


Public Function GetCellInfo(ByVal AddressOrRowIndex As Variant, Optional ByVal ColIndex As Long) As CellInfo
    Dim r As Long, c As Long
    Dim cellAddress As String, head As String, tail As String
    If VarType(AddressOrRowIndex) = vbString Then
        cellAddress = CStr(AddressOrRowIndex)
        If Len(cellAddress) < 2 Or Len(cellAddress) > 3 Then
            Err.Raise vbObjectError + 1331, _
                Source:="GridModel::GetCellInfo", _
                Description:="The cell address is invalid"
        End If
        head = UCase(Left(cellAddress, 1))
        tail = Mid(cellAddress, 2)
        If head < "A" Or head > "Z" Then
            Err.Raise vbObjectError + 1331, _
                Source:="GridModel::GetCellInfo", _
                Description:="The column label must be a letter between A and Z"
        End If
        c = Asc(head) - 65
        If Not IsNumeric(tail) Then
            Err.Raise vbObjectError + 1331, _
                Source:="GridModel::GetCellInfo", _
                Description:="The row index must be a number"
        End If
        r = CLng(tail)
        If r < 0 Or r > 99 Then
            Err.Raise vbObjectError + 1331, _
                Source:="GridModel::GetCellInfo", _
                Description:="The column index must be between 0 and 99"
        End If
    ElseIf VarType(AddressOrRowIndex) = vbInteger Or VarType(AddressOrRowIndex) = vbLong Then
        If IsMissing(ColIndex) Then
            Err.Raise vbObjectError + 1330, _
                Source:="GridModel::GetCellInfo", _
                Description:="A cell may be specified either by an address of the form ""X99"", or by a row and a column index"
        End If
        r = CLng(AddressOrRowIndex)
        If r < 0 Or r > 99 Then
            Err.Raise vbObjectError + 1331, _
                Source:="GridModel::GetCellInfo", _
                Description:="The row index must between 0 and 99"
        End If
        c = ColIndex
        If c < 0 Or c > 25 Then
            Err.Raise vbObjectError + 1331, _
                Source:="GridModel::GetCellInfo", _
                Description:="The column index must be between 0 and 25"
        End If
    End If
    Set GetCellInfo = FindCell(r, c)
End Function


Private Function FindCell(ByVal r As Long, c As Long) As CellInfo
    Dim i As Integer
    For i = 1 To m_Rows.Count
        If CLng(m_Rows(i)) = r Then
            If CLng(m_Cols(i)) = c Then
                Set FindCell = m_Cells(i)
                Exit Function
            End If
        End If
    Next
    ' Nothing found; create a new cell
    Dim cl As New CellInfo
    Set cl.Model = Me
    m_Rows.Add r
    m_Cols.Add c
    m_Cells.Add cl
    Set FindCell = cl
End Function


Public Sub Recalculate()
    ' Build the dependency graph
    Dim i As Integer, r As Integer, c As Integer, ci As CellInfo
    Dim cellAddress As String, v As Variant
    Dim G As New Graph
    For i = 1 To m_Rows.Count
        r = m_Rows(i)
        c = m_Cols(i)
        cellAddress = Chr(65 + c) & CStr(r)
        Set ci = m_Cells(i)
        If ci.CellType = Formula Then
            If Not G.NodeExists(cellAddress) Then G.AddNode cellAddress
            For Each v In ci.Dependencies
                G.AddEdge cellAddress, CStr(v)
            Next
        End If
    Next
    Dim cl As Collection
    Set cl = G.TopologicalSort()
    For Each v In cl
        Set ci = GetCellInfo(CStr(v))
        ci.Recalculate
    Next
End Sub


Public Function StringToVariant(ByVal StringValue As String) As Variant
    If Len(StringValue) = 0 Or UCase(StringValue) = "NULL" Then
        StringToVariant = vbNull
    ElseIf UCase(StringValue) = "TRUE" Or UCase(StringValue) = "FALSE" Then
        StringToVariant = CBool(StringValue)
    ElseIf IsNumeric(StringValue) Then
        StringToVariant = CSng(StringValue)
    Else
        StringToVariant = StringValue
    End If
End Function


Public Function GetAllCells() As Collection
    Dim result As New Collection
    Dim i As Integer, r As Integer, c As Integer, ci As CellInfo
    For i = 1 To m_Rows.Count
        r = m_Rows(i)
        c = m_Cols(i)
        Set ci = m_Cells(i)
        result.Add Array(r, c, ci.CellValue)
    Next
    Set GetAllCells = result
End Function
