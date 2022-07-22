VERSION 5.00
Begin VB.Form CRUDForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CRUD"
   ClientHeight    =   2496
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   5052
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
   ScaleHeight     =   2496
   ScaleWidth      =   5052
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox L 
      Height          =   1572
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1932
   End
   Begin VB.TextBox Tp 
      Height          =   288
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   972
   End
   Begin VB.Label Lp 
      AutoSize        =   -1  'True
      Caption         =   "&Filter prefix:"
      Height          =   192
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   840
   End
End
Attribute VB_Name = "CRUDForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_People As People

Public Property Get DataSource() As People
    Set DataSource = m_People
End Property

Public Property Set DataSource(ByRef value As People)
    Set m_People = value
End Property

Private Sub Form_Load()
    FillList
End Sub

Private Sub FillList(Optional ByVal filterPrefix As String = "")
    With L
        .Clear
        Dim Filter As SurnameFilter
        If Len(filterPrefix) > 0 Then
            Set Filter = New SurnameFilter
            Filter.SurnamePrefix = filterPrefix
        End If
        Dim i As PersonIterator
        Set i = DataSource.GetIterator(Filter)
        Do Until i.EOF
            .AddItem i.Current.Surname & ", " & i.Current.Name
            i.MoveNext
        Loop
    End With
End Sub

Private Sub Tp_Change()
    FillList Tp.Text
End Sub
