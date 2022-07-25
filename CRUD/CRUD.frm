VERSION 5.00
Begin VB.Form CRUDForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CRUD"
   ClientHeight    =   2592
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
   ScaleHeight     =   2592
   ScaleWidth      =   5052
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Bclose 
      Cancel          =   -1  'True
      Caption         =   "C&lose"
      Height          =   372
      Left            =   3960
      TabIndex        =   10
      Top             =   2160
      Width           =   972
   End
   Begin VB.TextBox Ts 
      Height          =   288
      Left            =   3000
      TabIndex        =   6
      Top             =   960
      Width           =   1932
   End
   Begin VB.TextBox Tn 
      Height          =   288
      Left            =   3000
      TabIndex        =   4
      Top             =   480
      Width           =   1932
   End
   Begin VB.CommandButton Bd 
      Caption         =   "&Delete"
      Enabled         =   0   'False
      Height          =   372
      Left            =   2280
      TabIndex        =   9
      Top             =   2160
      Width           =   972
   End
   Begin VB.CommandButton Bu 
      Caption         =   "&Update"
      Enabled         =   0   'False
      Height          =   372
      Left            =   1200
      TabIndex        =   8
      Top             =   2160
      Width           =   972
   End
   Begin VB.CommandButton Bc 
      Caption         =   "&Create"
      Height          =   372
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   972
   End
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
   Begin VB.Label Ls 
      AutoSize        =   -1  'True
      Caption         =   "&Surname:"
      Height          =   192
      Left            =   2160
      TabIndex        =   5
      Top             =   1008
      Width           =   684
   End
   Begin VB.Label Ln 
      AutoSize        =   -1  'True
      Caption         =   "&Name:"
      Height          =   192
      Left            =   2160
      TabIndex        =   3
      Top             =   528
      Width           =   456
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
Private m_IsDirty As Boolean
Private m_FormIsUpdating As Boolean
Private m_Id As Integer

Public Property Get DataSource() As People
    Set DataSource = m_People
End Property


Public Property Set DataSource(ByRef value As People)
    Set m_People = value
End Property


Public Property Get IsDirty() As Boolean
    IsDirty = m_IsDirty
End Property


Public Property Let IsDirty(ByVal value As Boolean)
    m_IsDirty = value
End Property


Public Property Get FormIsUpdating() As Boolean
    FormIsUpdating = m_FormIsUpdating
End Property


Public Property Let FormIsUpdating(ByVal value As Boolean)
    m_FormIsUpdating = value
End Property


Public Property Get id() As Integer
    id = m_Id
End Property


Public Property Let id(ByVal value As Integer)
    m_Id = value
End Property


Private Sub Bc_Click()
    DoCreate
End Sub

Private Sub Bclose_Click()
    ' This is not part of the spec, but I want a Close button, dammit!
    ' It's a bit academic makig this check as changes aren't persistent, but
    ' this could change
    If CanProceedAfterCheckingForChanges() Then
        Unload Me
    End If
End Sub

Private Sub Bd_Click()
    DoDelete
End Sub

Private Sub Bu_Click()
    DoUpdate
End Sub

Private Sub Form_Load()
    FillList Tp.Text
    ClearForm
End Sub


Private Sub FillList(Optional ByVal filterPrefix As String = "")
    With L
        .Clear
        ' Note: This is not in the given specification, but it is difficult to
        ' get back to an empty form to create a new record otherwise (unless
        ' you use hacks like setting a filter and clearing it)
        .AddItem "(new person)"
        .ItemData(.NewIndex) = -1
        Dim Filter As SurnameFilter
        If Len(filterPrefix) > 0 Then
            Set Filter = New SurnameFilter
            Filter.SurnamePrefix = filterPrefix
        End If
        Dim i As PersonIterator
        Set i = DataSource.GetIterator(Filter)
        If Not (i.BOF And i.EOF) Then
            Do Until i.EOF
                .AddItem i.Current.Surname & ", " & i.Current.Name
                .ItemData(.NewIndex) = i.Current.id
                i.MoveNext
            Loop
        End If
    End With
End Sub


Private Sub L_Click()
    Dim selectedId As Integer
    selectedId = L.ItemData(L.ListIndex)
    If selectedId = -1 Then
        If CanProceedAfterCheckingForChanges() Then
            ClearForm
        End If
    Else
        LoadPerson selectedId
        UpdateForm
    End If
End Sub


Private Sub Tn_Change()
    If FormIsUpdating Then Exit Sub
    IsDirty = True
    UpdateForm
End Sub


Private Sub Tn_GotFocus()
    SelectTextOnKeyboadFocus Tn
End Sub


Private Sub Tp_Change()
    FillList Tp.Text
End Sub

Private Sub LoadPerson(ByVal id As Integer)
    If Not CanProceedAfterCheckingForChanges() Then
        Exit Sub
    End If
    ClearForm updateAfter:=False
    FormIsUpdating = True
    Dim p As Person
    Set p = DataSource.GetPersonByID(id)
    Me.id = p.id
    Tn.Text = p.Name
    Ts.Text = p.Surname
    FormIsUpdating = False
    UpdateForm
End Sub


Private Sub Tp_GotFocus()
    SelectTextOnKeyboadFocus Tp
End Sub


Private Sub Tp_KeyUp(KeyCode As Integer, Shift As Integer)
    ' Allow navigation to the list by pressing the Down arrow
    If Shift = 0 And KeyCode = vbKeyDown Then
        L.SetFocus
    End If
End Sub

Private Sub Ts_Change()
    If FormIsUpdating Then Exit Sub
    IsDirty = True
    UpdateForm
End Sub


Private Sub UpdateForm()
    Bc.Default = False
    Bu.Default = False
    Bc.Enabled = (Me.id = -1 And (Len(Tn.Text) > 0 Or Len(Ts.Text) > 0))
    Bu.Enabled = (Me.id <> -1) And Me.IsDirty
    Bd.Enabled = (Me.id <> -1)
    Bc.Default = Bc.Enabled
    Bu.Default = Bu.Enabled
End Sub


Private Sub ClearForm(Optional ByVal updateAfter As Boolean = True)
    FormIsUpdating = True
    id = -1
    Tn.Text = ""
    Ts.Text = ""
    IsDirty = False
    FormIsUpdating = False
    If updateAfter Then UpdateForm
End Sub


Public Sub DoCreate()
    Dim newPerson As New Person
    With newPerson
        .Name = Trim(Tn.Text)
        .Surname = Trim(Ts.Text)
    End With
    DataSource.Add newPerson
    ClearForm
    FillList Tp.Text
    Tn.SetFocus
End Sub


Public Sub DoUpdate()
    Dim updatedPerson As New Person
    With updatedPerson
        .id = Me.id
        .Name = Trim(Tn.Text)
        .Surname = Trim(Ts.Text)
    End With
    DataSource.Update updatedPerson
    ClearForm
    FillList Tp.Text
    Tn.SetFocus
End Sub


'* \brief Selects the text in the control if the focus was moved there by
'* pressing Tab or using a Alt+character key, to be consistent with
'* normal dialog box behaviour
Private Sub SelectTextOnKeyboadFocus(ByRef textBox As textBox)
    Dim tabState As Integer: tabState = GetKeyState(VK_TAB)
    Dim altState As Integer: altState = GetKeyState(VK_LMENU)
    If tabState <> 0 Or altState <> 0 Then
        textBox.SelStart = 0
        textBox.SelLength = Len(textBox.Text)
    End If
End Sub


Private Sub Ts_GotFocus()
    SelectTextOnKeyboadFocus Ts
End Sub


Private Function PromptForSave() As Integer
    PromptForSave = MsgBox _
    ( _
        "Do you want to save the current record?", _
        Buttons:=vbInformation Or vbYesNoCancel, _
        Title:=Me.Caption _
    )
End Function


Private Function CanProceedAfterCheckingForChanges() As Boolean
    If Not IsDirty Then
        CanProceedAfterCheckingForChanges = True
        Exit Function
    End If
    Dim response As Integer
    response = PromptForSave()
    If response = vbCancel Then
        CanProceedAfterCheckingForChanges = False
        Exit Function
    End If
    If response = vbYes Then
        If id = -1 Then
            DoCreate
        Else
            DoUpdate
        End If
        CanProceedAfterCheckingForChanges = False
    End If
End Function


Public Sub DoDelete()
    If PromptForDelete() = vbYes Then
        Me.DataSource.Delete Me.id
        ClearForm
        FillList Tp.Text
        Tp.SetFocus
    End If
End Sub


Private Function PromptForDelete()
    PromptForDelete = MsgBox _
    ( _
        "Are you sure you want to delete the selected person?", _
        Buttons:=vbYesNo Or vbInformation, _
        Title:=Me.Caption _
    )
End Function
