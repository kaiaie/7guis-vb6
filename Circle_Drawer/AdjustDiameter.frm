VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form AdjustDiameterForm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Adjust Diameter"
   ClientHeight    =   996
   ClientLeft      =   36
   ClientTop       =   276
   ClientWidth     =   3744
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   996
   ScaleWidth      =   3744
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.Slider S 
      Height          =   372
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   3732
      _ExtentX        =   6583
      _ExtentY        =   656
      _Version        =   393216
      Min             =   1
      Max             =   200
      SelStart        =   1
      TickStyle       =   3
      Value           =   1
   End
   Begin VB.Label L 
      Alignment       =   2  'Center
      Height          =   372
      Left            =   120
      TabIndex        =   0
      Tag             =   "Adjust diameter of circle at ({0},{1})."
      Top             =   120
      Width           =   3492
   End
End
Attribute VB_Name = "AdjustDiameterForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_TargetForm As DrawCircleForm
Private m_CircleIndex As Integer
Private m_CircleX As Single
Private m_CircleY As Single


Public Property Get TargetForm() As DrawCircleForm
    Set TargetForm = m_TargetForm
End Property


Public Property Set TargetForm(ByRef Value As DrawCircleForm)
    Set m_TargetForm = Value
End Property


Public Property Get CircleIndex() As Integer
    CircleIndex = m_CircleIndex
End Property


Public Property Let CircleIndex(ByVal Value As Integer)
    m_CircleIndex = Value
End Property


Public Property Get CircleX() As Single
    CircleX = m_CircleX
End Property


Public Property Let CircleX(ByVal Value As Single)
    m_CircleX = Value
End Property


Public Property Get CircleY() As Single
    CircleY = m_CircleY
End Property


Public Property Let CircleY(ByRef Value As Single)
    m_CircleY = Value
End Property


Private Sub Form_Load()
    If Not TargetForm Is Nothing Then
        S.Value = TargetForm.S(CircleIndex).Width
    End If
    RefreshForm
End Sub


Public Sub RefreshForm()
    L.Caption = Strings.Format _
    ( _
        L.Tag, _
        CircleX, _
        CircleY _
    )
End Sub

Private Sub S_Scroll()
    If TargetForm Is Nothing Then Exit Sub
    TargetForm.AdjustDiameter CircleIndex, S.Value
End Sub
