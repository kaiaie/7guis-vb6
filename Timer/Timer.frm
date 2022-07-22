VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form TimerForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Timer"
   ClientHeight    =   1800
   ClientLeft      =   36
   ClientTop       =   324
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
   ScaleHeight     =   1800
   ScaleWidth      =   3744
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer T 
      Interval        =   100
      Left            =   120
      Top             =   1800
   End
   Begin VB.CommandButton R 
      Caption         =   "&Reset"
      Height          =   372
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   3492
   End
   Begin MSComctlLib.Slider S 
      Height          =   372
      Left            =   840
      TabIndex        =   4
      Top             =   840
      Width           =   2772
      _ExtentX        =   4890
      _ExtentY        =   656
      _Version        =   393216
      LargeChange     =   10
      Min             =   1
      Max             =   600
      SelStart        =   200
      TickStyle       =   3
      Value           =   200
   End
   Begin MSComctlLib.ProgressBar G 
      Height          =   252
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   2412
      _ExtentX        =   4255
      _ExtentY        =   445
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Ls 
      AutoSize        =   -1  'True
      Caption         =   "&Duration:"
      Height          =   192
      Left            =   120
      TabIndex        =   3
      Top             =   930
      Width           =   636
   End
   Begin VB.Label Le 
      AutoSize        =   -1  'True
      Caption         =   "0.0s"
      Height          =   192
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   288
   End
   Begin VB.Label Lg 
      AutoSize        =   -1  'True
      Caption         =   "Elapsed Time:"
      Height          =   192
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   984
   End
End
Attribute VB_Name = "TimerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_ElapsedTicks As Long
Private m_IsUpdating As Boolean


Private Sub R_Click()
    m_ElapsedTicks = 0
End Sub


Private Sub S_Scroll()
    UpdateForm
End Sub


Private Sub T_Timer()
    If m_ElapsedTicks < S.Value Then
        m_ElapsedTicks = m_ElapsedTicks + 1
        UpdateForm
    End If
End Sub


Private Sub UpdateForm()
    If m_IsUpdating Then Exit Sub
    m_IsUpdating = True
    Le.Caption = Format(m_ElapsedTicks / 10, "#0.0") & "s"
    G.Value = Clamp((m_ElapsedTicks / S.Value) * G.Max, G.Max)
    m_IsUpdating = False
End Sub


Private Function Clamp(ByVal currentValue As Single, ByVal maxValue As Single) As Single
    If currentValue > maxValue Then
        Clamp = maxValue
    Else
        Clamp = currentValue
    End If
End Function
