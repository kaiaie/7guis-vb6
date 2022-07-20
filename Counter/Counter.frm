VERSION 5.00
Begin VB.Form CounterForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Counter"
   ClientHeight    =   540
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   2436
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
   ScaleHeight     =   540
   ScaleWidth      =   2436
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton B 
      Caption         =   "&Count"
      Default         =   -1  'True
      Height          =   288
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   972
   End
   Begin VB.TextBox T 
      Height          =   288
      Left            =   120
      TabIndex        =   0
      Text            =   "0"
      Top             =   120
      Width           =   972
   End
End
Attribute VB_Name = "CounterForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const EM_SETREADONLY = &HCF

Private Sub B_Click()
    T.Text = CStr(CLng(T.Text) + 1)
End Sub

Private Sub Form_Load()
    SendMessage T.hwnd, EM_SETREADONLY, 1&, 0&
End Sub
