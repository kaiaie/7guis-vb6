VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form CellsForm 
   Caption         =   "Cells"
   ClientHeight    =   4776
   ClientLeft      =   132
   ClientTop       =   708
   ClientWidth     =   7704
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4776
   ScaleWidth      =   7704
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox T 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   0
      TabIndex        =   1
      Top             =   4440
      Visible         =   0   'False
      Width           =   612
   End
   Begin MSFlexGridLib.MSFlexGrid G 
      Height          =   4452
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7572
      _ExtentX        =   13356
      _ExtentY        =   7853
      _Version        =   393216
      Rows            =   101
      Cols            =   27
      AllowUserResizing=   3
   End
   Begin VB.Menu FileMenu 
      Caption         =   "&File"
      Begin VB.Menu FileExitMenuItem 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "CellsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Model As New GridModel

Private Sub FileExitMenuItem_Click()
    End
End Sub

Private Sub Form_Load()
    SetGridLabels
    G.Row = 1
    G.Col = 1
End Sub


Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    ' Resize grid control to fill form window
    G.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub


Private Sub SetGridLabels()
    Dim i As Integer
    ' Row labels
    G.Col = 0
    For i = 1 To 100
        G.Row = i
        G.Text = CStr(i - 1)
        G.CellFontBold = True
    Next
    
    ' Column labels
    G.Row = 0
    For i = 1 To 26
        G.Col = i
        G.Text = CStr(Chr(64 + i))
        G.CellFontBold = True
        G.CellAlignment = flexAlignCenterBottom
    Next
End Sub

Private Sub G_DblClick()
    EditCell
End Sub

Private Sub G_KeyUp(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyF2 Then
        EditCell
    End If
End Sub


Private Sub EditCell()
    With T
        .Move G.ColPos(G.Col), G.RowPos(G.Row), G.ColWidth(G.Col)
        .Visible = True
        .Text = Model.GetCellInfo(G.Row - G.FixedRows, G.Col - G.FixedCols).CellValue
        .SetFocus
    End With
End Sub


Private Sub DiscardCellEdits()
    With T
        .Visible = False
    End With
    G.SetFocus
End Sub


Private Sub AcceptCellEdits()
    Dim cl As CellInfo
    Set cl = Model.GetCellInfo(G.Row - G.FixedRows, G.Col - G.FixedCols)
    With T
        .Visible = False
        If Left(.Text, 1) = "=" Then
            cl.CellType = Formula
            cl.CellFormula = Mid(.Text, 2)
        Else
            cl.CellType = Literal
            cl.CellValue = Model.StringToVariant(.Text)
        End If
    End With
    G.SetFocus
    RefreshGrid
End Sub


Private Sub T_KeyUp(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then
        DiscardCellEdits
    ElseIf Shift = 0 And KeyCode = vbKeyReturn Then
        AcceptCellEdits
        ' Move to the next row
        If G.Row < 100 Then
            G.Row = G.Row + 1
        End If
    End If
End Sub


Private Sub T_LostFocus()
    DiscardCellEdits
End Sub


Private Function FormatCell(ByVal CellValue As Variant) As String
    If IsNull(CellValue) Then
        FormatCell = "#NULL"
    Else
        FormatCell = CStr(CellValue)
    End If
End Function


Private Function RefreshGrid()
    Dim errorOccurred As Boolean
TryRecalculate:
    On Error GoTo ErrRecalculate
    Model.Recalculate
    GoTo EndRecalculate

ErrRecalculate:
    If Err.Number = (vbObjectError + 990) Then
        MsgBox _
            Prompt:="You have a circular reference somewhere in your formulas!", _
            Buttons:=vbInformation Or vbOKOnly, _
            Title:=Me.Caption
    Else
        MsgBox _
            Prompt:=Strings.Format( _
                "An unexpected error ""{0}"" (code: {1}) occurred at {2}", _
                Err.Description, Err.Number, Err.Source), _
            Buttons:=vbCritical Or vbOKOnly, _
            Title:=Me.Caption
    End If
    errorOccurred = True

EndRecalculate:
    On Error GoTo 0
    
    If Not errorOccurred Then
        Dim savedRow As Long, savedCol As Long
        savedRow = G.Row: savedCol = G.Col
        G.Clear
        SetGridLabels
        Dim cell As Variant
        Dim r As Integer, c As Integer, v As Variant
        For Each cell In Model.GetAllCells()
            G.Row = cell(0) + G.FixedRows
            G.Col = cell(1) + G.FixedCols
            G.Text = FormatCell(cell(2))
        Next
        G.Row = savedRow: G.Col = savedCol
    End If
    Exit Function
End Function

