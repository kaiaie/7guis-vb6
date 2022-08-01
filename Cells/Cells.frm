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
      Left            =   360
      TabIndex        =   1
      Top             =   2640
      Visible         =   0   'False
      Width           =   612
   End
   Begin MSFlexGridLib.MSFlexGrid G 
      Height          =   2412
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3732
      _ExtentX        =   6583
      _ExtentY        =   4255
      _Version        =   393216
      Rows            =   101
      Cols            =   27
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
        .Text = G.Text
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
    Set cl = Model.GetCellInfo(G.Row, G.Col)
    With T
        .Visible = False
        If Left(.Text, 1) = "=" Then
            cl.CellType = Formula
        Else
            cl.CellType = Literal
            If IsNumeric(.Text) Then
            Else
            End If
        End If
        G.Text = .Text
    End With
    G.SetFocus
End Sub


Private Sub T_KeyUp(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then
        DiscardCellEdits
    ElseIf Shift = 0 And KeyCode = vbKeyReturn Then
        AcceptCellEdits
    End If
End Sub


Private Sub T_LostFocus()
    DiscardCellEdits
End Sub

