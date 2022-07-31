VERSION 5.00
Begin VB.Form DrawCircleForm 
   Caption         =   "CircleDraw"
   ClientHeight    =   4008
   ClientLeft      =   132
   ClientTop       =   420
   ClientWidth     =   5676
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4008
   ScaleWidth      =   5676
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Br 
      Caption         =   "&Redo"
      Enabled         =   0   'False
      Height          =   372
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   972
   End
   Begin VB.CommandButton Bu 
      Caption         =   "&Undo"
      Enabled         =   0   'False
      Height          =   372
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   972
   End
   Begin VB.PictureBox P 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000007&
      Height          =   3372
      Left            =   0
      ScaleHeight     =   277
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   467
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   600
      Width           =   5652
      Begin VB.Shape S 
         BackStyle       =   1  'Opaque
         Height          =   612
         Index           =   0
         Left            =   240
         Shape           =   3  'Circle
         Top             =   2760
         Visible         =   0   'False
         Width           =   612
      End
   End
   Begin VB.Menu CirclePopupMenu 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu AdjustDiameterMenuItem 
         Caption         =   "&Adjust Diameter"
      End
   End
End
Attribute VB_Name = "DrawCircleForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Stack As UndoStack
Private m_CircleCount As Integer
Private m_SelectedCircle As Shape


Public Property Get CircleCount() As Integer
    CircleCount = m_CircleCount
End Property


Private Property Let CircleCount(ByVal Value As Integer)
    m_CircleCount = Value
End Property


Private Property Get SelectedCircle() As Shape
    Set SelectedCircle = m_SelectedCircle
End Property


Private Property Set SelectedCircle(ByRef Value As Shape)
    Set m_SelectedCircle = Value
    RefreshForm
End Property


Public Property Get Stack() As UndoStack
    Set Stack = m_Stack
End Property


Private Property Set Stack(ByRef Value As UndoStack)
    Set m_Stack = Value
End Property


Private Sub AdjustDiameterMenuItem_Click()
    ShowAdjustmentForm
End Sub

Private Sub Br_Click()
    If Stack.CanRedo Then Stack.Redo
    RefreshForm
End Sub

Private Sub Bu_Click()
    If Stack.CanUndo Then Stack.Undo
    RefreshForm
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    ' Shortcut keys for Undo/Redo
    If (Shift And vbCtrlMask) = vbCtrlMask And KeyCode = vbKeyZ Then
        If Stack.CanUndo Then Stack.Undo
    ElseIf (Shift And vbCtrlMask) = vbCtrlMask And KeyCode = vbKeyY Then
        If Stack.CanRedo Then Stack.Redo
    End If
End Sub

Private Sub Form_Load()
    Set Stack = New UndoStack
    Set Stack.OwnerForm = Me
End Sub


Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    
    ' Centre buttons at top of form
    Dim w As Single
    w = (Br.Left + Br.Width) - Bu.Left
    If w < Me.ScaleWidth Then
        Dim L As Single
        L = (Me.ScaleWidth - w) / 2
        Bu.Move L
        L = L + Bu.Width + (8# * Screen.TwipsPerPixelX)
        Br.Move L
    End If
    
    ' Resize picture box to remainder of form
    P.Move 0, P.Top, Me.ScaleWidth, Me.ScaleHeight - P.Top
End Sub


Private Sub RefreshForm()
    ' Highlight the selected circle (if one is selected)
    If Not m_SelectedCircle Is Nothing Then
        Dim ctl As Control
        For Each ctl In Me.Controls
            If ctl.Name = "S" And ctl.Visible And TypeOf ctl Is Shape Then
                Dim crcl As Shape: Set crcl = ctl
                If crcl Is m_SelectedCircle Then
                    crcl.BackColor = vbButtonFace
                    crcl.BorderColor = vbButtonText
                Else
                    crcl.BackColor = vbWindowBackground
                    crcl.BorderColor = vbWindowText
                End If
            End If
        Next
    End If
    
    ' Enable or diable Undo/Redo buttons based on state of undo stack
    Bu.Enabled = Stack.CanUndo
    Br.Enabled = Stack.CanRedo
    DoEvents
End Sub


Private Sub P_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim crcl As Shape
    Set crcl = GetCircleAtLocation(X, Y)
    If crcl Is Nothing Then
        If (Button And vbLeftButton) = vbLeftButton Then
            AddCircle CircleCount, X, Y
            Stack.AddCircle CircleCount, X, Y
            CircleCount = CircleCount + 1
        End If
        RefreshForm
    Else
        Set SelectedCircle = crcl
        If (Button And vbRightButton) = vbRightButton Then
            Me.PopupMenu CirclePopupMenu, _
                X:=X * Screen.TwipsPerPixelX + P.Left, _
                Y:=Y * Screen.TwipsPerPixelY + P.Top
        End If
    End If
End Sub


Private Function GetCircleAtLocation(ByVal X As Single, ByVal Y As Single) As Shape
    Dim result As Shape
    Set result = Nothing
    Dim ctl As Control
    For Each ctl In Me.Controls
        If ctl.Name = "S" And ctl.Visible And TypeOf ctl Is Shape Then
            Dim crcl As Shape: Set crcl = ctl
            Dim cX As Single, cY As Single, cR As Single
            cR = crcl.Width / 2
            cX = crcl.Left + cR
            cY = crcl.Top + cR
            If Sqr((X - cX) * (X - cX) + (Y - cY) * (Y - cY)) <= cR Then
                Set result = crcl
                Exit For
            End If
        End If
    Next
    Set GetCircleAtLocation = result
End Function


Public Sub AdjustDiameter(ByVal Index As Integer, ByVal Diameter As Single)
    Dim X As Single, Y As Single, d As Single, r As Single
    Dim crcl As Shape: Set crcl = S(Index)
    X = crcl.Left + crcl.Width / 2
    Y = crcl.Top + crcl.Height / 2
    r = Diameter / 2
    crcl.Move X - r, Y - r, Diameter, Diameter
End Sub


Public Sub ShowAdjustmentForm()
    If SelectedCircle Is Nothing Then Exit Sub
    Dim OldDiameter As Single
    OldDiameter = SelectedCircle.Width
    Dim f As New AdjustDiameterForm
    With f
        Set .TargetForm = Me
        .CircleIndex = SelectedCircle.Index
        .CircleX = SelectedCircle.Left + (SelectedCircle.Width / 2)
        .CircleY = SelectedCircle.Top + (SelectedCircle.Height / 2)
    End With
    f.Show vbModal, OwnerForm:=Me
    Stack.AddDiameterAdjustment SelectedCircle.Index, OldDiameter, SelectedCircle.Width
    RefreshForm
End Sub


Public Sub AddCircle(ByVal Index As Integer, ByVal X As Single, ByVal Y As Single)
    Dim crcl As Shape
    If Index = 0 Then
        Set crcl = S(0)
    Else
        Load S(Index)
        Set crcl = S(Index)
    End If
    crcl.Visible = True
    crcl.Move X - (crcl.Width / 2), Y - (crcl.Height / 2)
    RefreshForm
End Sub


Public Sub RemoveCircle(ByVal Index As Integer)
    If Index = 0 Then
        S(0).Visible = False
    Else
        Unload S(Index)
    End If
    RefreshForm
End Sub
