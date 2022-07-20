VERSION 5.00
Begin VB.Form ConverterForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Temperature Converter"
   ClientHeight    =   552
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   3852
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
   ScaleHeight     =   552
   ScaleWidth      =   3852
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Tf 
      Height          =   288
      Left            =   1920
      TabIndex        =   3
      Top             =   120
      Width           =   972
   End
   Begin VB.TextBox Tc 
      Height          =   288
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   972
   End
   Begin VB.Label Lf 
      AutoSize        =   -1  'True
      Caption         =   "°&Fahrenheit"
      Height          =   192
      Left            =   3000
      TabIndex        =   2
      Top             =   168
      Width           =   804
   End
   Begin VB.Label Lc 
      AutoSize        =   -1  'True
      Caption         =   "°&Celsius ="
      Height          =   192
      Left            =   1200
      TabIndex        =   0
      Top             =   168
      Width           =   708
   End
End
Attribute VB_Name = "ConverterForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Const VK_TAB = &H9
Private Const VK_LMENU = &HA4
Private Const GWL_STYLE = (-16)
Private Const ES_NUMBER = &H2000&

Private m_FormIsUpdating As Boolean

Private Function CelsiusToFahrenheit(ByVal celsius As Double) As Double
    CelsiusToFahrenheit = (celsius * 1.8) + 32
End Function

Private Function FahrenheitToCelsius(ByVal fahrenheit As Double) As Double
    FahrenheitToCelsius = (fahrenheit - 32) / 1.8
End Function

Private Sub SetNumericOnlyInput(ByRef textBox As textBox)
    Dim style As Long
    style = GetWindowLong(textBox.hwnd, GWL_STYLE)
    SetWindowLong textBox.hwnd, GWL_STYLE, style + ES_NUMBER
End Sub

Private Sub SelectTextOnKeyboadFocus(ByRef textBox As textBox)
    Dim tabState As Integer: tabState = GetKeyState(VK_TAB)
    Dim altState As Integer: altState = GetKeyState(VK_LMENU)
    If tabState <> 0 Or altState <> 0 Then
        textBox.SelStart = 0
        textBox.SelLength = Len(textBox.Text)
    End If
End Sub

Private Sub Form_Load()
    SetNumericOnlyInput Tc
    SetNumericOnlyInput Tf
End Sub

Private Sub Tc_Change()
    If m_FormIsUpdating Then Exit Sub
    m_FormIsUpdating = True
    If Tc.Text = "" Or Not IsNumeric(Tc.Text) Then
        Tf.Text = ""
    Else
        Tf.Text = CStr(CLng(CelsiusToFahrenheit(CDbl(Tc.Text))))
    End If
    m_FormIsUpdating = False
End Sub

Private Sub Tc_GotFocus()
    SelectTextOnKeyboadFocus Tc
End Sub

Private Sub Tf_Change()
    If m_FormIsUpdating Then Exit Sub
    m_FormIsUpdating = True
    If Tf.Text = "" Then
        Tc.Text = ""
    Else
        Tc.Text = CStr(CLng(FahrenheitToCelsius(CDbl(Tf.Text))))
    End If
    m_FormIsUpdating = False
End Sub

Private Sub Tf_GotFocus()
    SelectTextOnKeyboadFocus Tf
End Sub
