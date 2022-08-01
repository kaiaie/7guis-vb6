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


Public Function GetCellInfo(ByVal AddressOrRowIndex As Variant, Optional ByVal ColIndex As Integer) As CellInfo
    Dim R As Integer, Col As Integer
    Dim cellAddress As String, head As String, tail As String
    If VarType(AddressOrRowIndex) = vbString Then
        If Not IsMissing(ColIndex) Then
            Err.Raise vbObjectError + 1330, _
                Description:="A cell may be specified either by an address of the form ""X99"", or by a row and a column index"
        End If
        cellAddress = CStr(AddressOrRowIndex)
        If Len(cellAddress) < 2 Or Len(cellAddress) > 3 Then
            Err.Raise vbObjectError + 1331, _
                Description:="The cell address is invalid"
        End If
        head = UCase(Left(cellAddress, 1))
        tail = Mid(cellAddress, 2)
        If head < "A" Or head > "Z" Then
            Err.Raise vbObjectError + 1331, _
                Description:="The column label must be a letter between A and Z"
        End If
        R = Asc(head) - 64
        If Not IsNumeric(tail) Then
            Err.Raise vbObjectError + 1331, _
                Description:="The row index must be a number"
        End If
        c = CInt(tail)
        If c < 0 Or c > 99 Then
            Err.Raise vbObjectError + 1331, _
                Description:="The column index must be between 0 and 99"
        End If
    ElseIf VarType(AddressOrRowIndex) = vbInteger Then
        If IsMissing(ColIndex) Then
            Err.Raise vbObjectError + 1330, _
                Description:="A cell may be specified either by an address of the form ""X99"", or by a row and a column index"
        End If
        R = CInt(AddressOrRowIndex)
        If R < 1 Or R > 26 Then
            Err.Raise vbObjectError + 1331, _
                Description:="The column label must between 1 and 26"
        End If
        c = ColIndex
        If c < 0 Or c > 99 Then
            Err.Raise vbObjectError + 1331, _
                Description:="The column index must be between 0 and 99"
        End If
    End If
    Set GetCellInfo = FindCell(R, c)
End Function


Private Function FindCell(ByVal R As Integer, c As Integer) As CellInfo
    Dim i As Integer
    For i = 1 To m_Rows.Count
        If CInt(m_Rows(i)) = R Then
            If CInt(m_Cols(i)) = c Then
                Set FindCell = m_Cells(i)
                Exit Function
            End If
        End If
    Next
    ' Nothing found; create a new cell
    Dim cl As New CellInfo
    m_Rows.Add R
    m_Cols.Add c
    m_Cells.Add cl
    Set FindCell = cl
End Function


Public Sub Recalculate()

End Sub
